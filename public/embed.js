/* WebEmbed: on-slide logic for the PowerPoint content add-in.
 *
 * Each inserted add-in is its own instance with its own settings, so one deck
 * can hold many embeds. What gets stored (in the .pptx) is:
 *   { mode: "url",  url,  width }          a live page, fetched every time the slide shows
 *   { mode: "html", html, name, width }    inline HTML, rendered in a sandboxed iframe
 * width is the page width in CSS px that gets scaled to fit the box (0 = box width).
 *
 * Outside PowerPoint (plain browser preview) the same code runs against
 * localStorage, so the UI can be developed without Office.
 */
(() => {
  "use strict";

  const KEY = "webembed";
  // Office keeps add-in settings in the document and limits their size
  // (documented as 2 MB per add-in). Stay well under it.
  const MAX_INLINE_BYTES = 1_500_000;

  const SAMPLES = [
    { file: "samples/orbit.html", title: "Orbit", note: "Particle orbit loop, canvas" },
    { file: "samples/kinetic.html", title: "Kinetic type", note: "Words that cycle in and out" },
    { file: "samples/bars.html", title: "Bar chart build", note: "Animated SVG chart, replays" },
  ];

  const URL_HINT = "The page loads live, so you need a connection when presenting.";

  const $ = (sel) => document.querySelector(sel);
  const stage = $("#stage");
  const config = $("#config");
  const bar = $("#bar");
  const barLabel = $("#bar-label");
  const urlInput = $("#url-input");
  const urlStatus = $("#url-status");
  const htmlInput = $("#html-input");
  const htmlSize = $("#html-size");
  const widthSel = $("#page-width");
  const applyBtn = $("#btn-apply");
  const errorEl = $("#error");
  const log = window.__wlog || (() => {});

  let inOffice = false;
  let state = null;
  let activeTab = "url";

  /* ---------------- storage ---------------- */

  // Older versions stored a zoom factor instead of a page width.
  function upgrade(s) {
    if (s && s.width == null) {
      const z = Number(s.zoom) || 1;
      s = { ...s, width: z === 1 ? 0 : Math.round(innerWidth / z) };
      delete s.zoom;
    }
    return s;
  }

  const store = {
    load() {
      if (inOffice) return upgrade(Office.context.document.settings.get(KEY) || null);
      try { return upgrade(JSON.parse(localStorage.getItem(KEY))); } catch { return null; }
    },
    save(value) {
      if (!inOffice) {
        try {
          value ? localStorage.setItem(KEY, JSON.stringify(value)) : localStorage.removeItem(KEY);
        } catch { /* preview only; ignore */ }
        return Promise.resolve();
      }
      return new Promise((resolve, reject) => {
        const s = Office.context.document.settings;
        value ? s.set(KEY, value) : s.remove(KEY);
        s.saveAsync((r) =>
          r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error));
      });
    },
  };

  /* ---------------- view mode (edit vs slideshow) ---------------- */

  function setView(view) {
    // PowerPoint reports "edit" for Normal view and "read" for Slide Show.
    document.body.dataset.view = view === "read" ? "read" : "edit";
    log("view", view);
    fit(); // the bar comes and goes, so the stage changes height
  }

  function watchView() {
    const doc = Office.context.document;
    if (typeof doc.getActiveViewAsync === "function") {
      doc.getActiveViewAsync((r) => { if (r.status === "succeeded") setView(r.value); });
    }
    if (Office.EventType && Office.EventType.ActiveViewChanged && doc.addHandlerAsync) {
      try { doc.addHandlerAsync(Office.EventType.ActiveViewChanged, (e) => setView(e.activeView)); }
      catch { /* not supported: the bar just stays in the slideshow */ }
    }
  }

  /* ---------------- URL helpers ---------------- */

  function normalizeUrl(raw) {
    let s = raw.trim();
    if (!s) throw new Error("Enter a web address.");
    if (!/^[a-z][a-z0-9+.-]*:/i.test(s)) s = "https://" + s;
    let u;
    try { u = new URL(s); } catch { throw new Error("That doesn't look like a web address."); }
    if (u.protocol !== "https:" && u.protocol !== "http:") throw new Error("Only http(s) addresses can be embedded.");
    if (u.protocol === "http:" && location.protocol === "https:") {
      throw new Error("PowerPoint blocks http:// pages inside the add-in. Use the https:// address.");
    }

    const host = u.hostname.replace(/^www\.|^m\./, "");
    // YouTube → embed player
    let yt = null;
    if (host === "youtube.com" && u.pathname === "/watch") yt = u.searchParams.get("v");
    else if (host === "youtube.com" && /^\/(shorts|live)\//.test(u.pathname)) yt = u.pathname.split("/")[2];
    else if (host === "youtu.be") yt = u.pathname.slice(1);
    if (yt) {
      const e = new URL("https://www.youtube.com/embed/" + encodeURIComponent(yt));
      const t = u.searchParams.get("t") || u.searchParams.get("start");
      if (t) e.searchParams.set("start", String(parseInt(t, 10) || 0));
      e.searchParams.set("rel", "0");
      return e.href;
    }
    // Vimeo → player
    const vm = host === "vimeo.com" && u.pathname.match(/^\/(\d+)/);
    if (vm) return "https://player.vimeo.com/video/" + vm[1];

    return u.href;
  }

  const hostOf = (url) => { try { return new URL(url).hostname.replace(/^www\./, ""); } catch { return url; } };

  // Asks the add-in server whether the site allows framing (X-Frame-Options / CSP).
  // The browser can't tell us this itself: a blocked frame still fires "load".
  // On static hosting there is no endpoint, and this returns null.
  async function checkFraming(url) {
    try {
      const r = await fetch("api/check?url=" + encodeURIComponent(url) + "&origin=" + encodeURIComponent(location.origin), { cache: "no-store" });
      if (!r.ok) return null;
      return await r.json();
    } catch {
      return null;
    }
  }

  /* ---------------- rendering ---------------- */

  // Lay the page out at the chosen width, then scale it to the box.
  function fit() {
    const f = stage.querySelector("iframe");
    if (!f || !state) return;
    const w = Number(state.width) || 0;
    const W = stage.clientWidth, H = stage.clientHeight;
    if (!w || !W) {
      f.style.width = f.style.height = f.style.transform = "";
      return;
    }
    const s = W / w;
    f.style.width = w + "px";
    f.style.height = H / s + "px";
    f.style.transform = `scale(${s})`;
  }

  // A page is shown once it fires "load" or after REVEAL_MS, whichever comes
  // first ("load" waits for every image and script, and most pages are
  // readable long before that). One that still hasn't loaded after
  // LOAD_TIMEOUT_MS gets a card on top; the page keeps loading underneath
  // and the card goes away if it arrives.
  const REVEAL_MS = 2000;
  const LOAD_TIMEOUT_MS = 15000;
  let timers = [];
  const clearTimers = () => { timers.forEach(clearTimeout); timers = []; };

  function render() {
    clearTimers();
    stage.replaceChildren();
    bar.hidden = !state;
    if (!state) return openConfig();

    const f = document.createElement("iframe");
    // Fade in. If the host isn't painting frames, the transition would sit at
    // its start (invisible), so drop it shortly after and jump to shown.
    const ready = () => {
      if (f.classList.contains("ready")) return;
      f.classList.add("ready");
      setTimeout(() => { f.style.transition = "none"; }, 600);
    };
    if (state.mode === "url") {
      const url = state.url;
      barLabel.textContent = hostOf(url);
      barLabel.title = url;
      // Only the author sees this; the slideshow stays empty until the page shows.
      const loading = document.createElement("div");
      loading.className = "loading edit-only";
      loading.textContent = `Loading ${hostOf(url)}…`;
      stage.append(loading);
      f.src = url;
      f.allow = "autoplay; fullscreen; encrypted-media; picture-in-picture; clipboard-write";
      f.referrerPolicy = "strict-origin-when-cross-origin";
      f.title = url;
      f.addEventListener("load", () => {
        clearTimers();
        loading.remove();
        stage.querySelector(".blocked.late")?.remove();
        ready();
        log("iframe-load", url);
      });
      timers.push(setTimeout(ready, REVEAL_MS));
      timers.push(setTimeout(() => {
        if (state && state.url === url) showUnreachable(url);
      }, LOAD_TIMEOUT_MS));
    } else {
      barLabel.textContent = state.name || "HTML";
      barLabel.title = "";
      // No allow-same-origin: inline HTML runs with an opaque origin and
      // cannot reach the add-in or Office APIs.
      f.setAttribute("sandbox", "allow-scripts allow-popups");
      f.allow = "autoplay; fullscreen";
      f.addEventListener("load", ready);
      f.srcdoc = state.html;
      f.title = state.name || "Embedded HTML";
    }
    stage.append(f);
    fit();

    // A site that refuses framing renders as a silent white box. Re-check on
    // every render (headers can change) and show a clear card instead.
    if (state.mode === "url") {
      const url = state.url;
      checkFraming(url).then((v) => {
        if (v && v.embeddable === false && state && state.url === url) showBlocked(url, v.reason);
      });
    }
  }

  function openInBrowser(url) {
    const ui = inOffice && Office.context.ui;
    if (ui && typeof ui.openBrowserWindow === "function") ui.openBrowserWindow(url);
    else window.open(url, "_blank", "noopener");
  }

  // overlay: put the card on top of the page instead of replacing it.
  function showCard(url, title, why, { retry = false, overlay = false } = {}) {
    clearTimers();
    const card = document.createElement("div");
    card.className = "blocked" + (overlay ? " late" : "");
    card.innerHTML = `
      <strong></strong>
      <p class="why"></p>
      <div class="actions">
        <button type="button" class="quiet" data-act="retry" hidden>Try again</button>
        <button type="button" class="primary" data-act="open">Open in browser</button>
      </div>`;
    card.querySelector("strong").textContent = title;
    card.querySelector(".why").textContent = why;
    card.querySelector('[data-act="open"]').onclick = () => openInBrowser(url);
    const again = card.querySelector('[data-act="retry"]');
    again.hidden = !retry;
    again.onclick = render;
    if (overlay) stage.append(card);
    else stage.replaceChildren(card);
  }

  function showBlocked(url, reason) {
    showCard(url, `${hostOf(url)} can't be shown inside a slide`, reason + " Only the site's owner can allow it.");
    log("blocked", `${url}: ${reason}`);
  }

  function showUnreachable(url) {
    stage.querySelector(".loading")?.remove();
    showCard(url, `${hostOf(url)} is still loading`, "", { retry: true, overlay: true });
    log("timeout", url);
  }

  /* ---------------- config panel ---------------- */

  function selectTab(tab) {
    activeTab = tab;
    for (const b of config.querySelectorAll("[role=tab]")) b.setAttribute("aria-selected", String(b.dataset.tab === tab));
    for (const p of config.querySelectorAll(".pane")) p.hidden = p.dataset.pane !== tab;
    applyBtn.hidden = tab === "samples";
    applyBtn.textContent = tab === "url" && urlInput.dataset.warned ? "Embed anyway" : "Embed";
    showError("");
    focusField();
  }

  function focusField() {
    const el = activeTab === "url" ? urlInput : activeTab === "html" ? htmlInput : null;
    if (el && !config.hidden) requestAnimationFrame(() => el.focus());
  }

  function resetUrlStatus() {
    delete urlInput.dataset.warned;
    applyBtn.textContent = "Embed";
    urlStatus.className = "hint";
    urlStatus.textContent = URL_HINT;
  }

  function openConfig() {
    $("#btn-cancel").hidden = !state;
    widthSel.value = String(state?.width ?? 0);
    if (!widthSel.value) widthSel.value = "0"; // a width that isn't in the list
    if (state?.mode === "url") urlInput.value = state.url;
    if (state?.mode === "html") htmlInput.value = state.html;
    resetUrlStatus();
    updateSize();
    config.hidden = false;
    bar.hidden = true;
    selectTab(state?.mode || activeTab);
  }

  function closeConfig() {
    config.hidden = true;
    bar.hidden = !state;
    showError("");
  }

  function showError(msg) {
    errorEl.textContent = msg;
    errorEl.hidden = !msg;
  }

  // Measure what is actually stored: the JSON-escaped string.
  const byteLength = (s) => new TextEncoder().encode(JSON.stringify(s)).length;

  function updateSize() {
    const n = htmlInput.value ? byteLength(htmlInput.value) : 0;
    if (!n) { htmlSize.textContent = ""; return; }
    const kb = (n / 1024).toFixed(n < 10240 ? 1 : 0);
    const over = n > MAX_INLINE_BYTES;
    htmlSize.textContent = over
      ? `${kb} KB: too big to store in the deck (limit ${Math.round(MAX_INLINE_BYTES / 1024)} KB). Host it and embed by URL instead.`
      : `${kb} KB`;
    htmlSize.className = "hint" + (over ? " warn" : "");
  }

  async function commit(next) {
    log("commit", next ? next.mode + " " + (next.url || next.name) : "clear");
    try {
      await store.save(next);
      log("saved", "ok");
    } catch (e) {
      log("save-failed", e && (e.message || e));
      showError("Couldn't save to the document: " + (e?.message || e));
      return;
    }
    state = next;
    closeConfig();
    render();
  }

  const chosenWidth = () => Number(widthSel.value) || 0;

  async function applyUrl() {
    if (applyBtn.disabled) return; // a framing check is already running
    let url;
    try { url = normalizeUrl(urlInput.value); } catch (e) { return showError(e.message); }
    showError("");
    urlInput.value = url;
    if (urlInput.dataset.warned !== url) {
      applyBtn.disabled = true;
      const verdict = await checkFraming(url);
      applyBtn.disabled = false;
      if (verdict && verdict.embeddable === false) {
        urlInput.dataset.warned = url;
        urlStatus.className = "hint warn";
        urlStatus.textContent = `${verdict.reason} The slide will show an Open in browser button instead of the page.`;
        applyBtn.textContent = "Embed anyway";
        return;
      }
    }
    await commit({ mode: "url", url, width: chosenWidth() });
  }

  async function applyHtml(html, name) {
    if (!html.trim()) return showError("Paste some HTML or open a file first.");
    if (byteLength(html) > MAX_INLINE_BYTES) {
      return showError("Too big to store in the deck. Host the file and embed it by URL instead.");
    }
    await commit({ mode: "html", html, name: name || "Inline HTML", width: chosenWidth() });
  }

  function buildSamples() {
    const list = $("#sample-list");
    for (const s of SAMPLES) {
      const li = document.createElement("li");
      li.innerHTML = `<span class="name"></span><button type="button">Use</button>`;
      li.querySelector(".name").innerHTML = `${s.title}<small>${s.note}</small>`;
      // Copied into the deck, so the sample doesn't depend on this server staying up.
      li.querySelector("button").onclick = async () => {
        try {
          const r = await fetch(s.file);
          if (!r.ok) throw new Error(r.status);
          await applyHtml(await r.text(), s.title);
        } catch (e) { showError("Couldn't load the sample: " + e.message); }
      };
      list.append(li);
    }
  }

  function wireUi() {
    for (const b of config.querySelectorAll("[role=tab]")) b.onclick = () => selectTab(b.dataset.tab);
    applyBtn.onclick = () => (activeTab === "url" ? applyUrl() : applyHtml(htmlInput.value));
    $("#btn-cancel").onclick = closeConfig;
    $("#btn-edit").onclick = openConfig;
    $("#btn-reload").onclick = render;
    urlInput.addEventListener("keydown", (e) => { if (e.key === "Enter") applyUrl(); });
    urlInput.addEventListener("input", resetUrlStatus);
    htmlInput.addEventListener("input", updateSize);
    htmlInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) applyHtml(htmlInput.value);
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && state && !config.hidden) closeConfig();
    });
    $("#html-file").addEventListener("change", async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      htmlInput.value = await file.text();
      updateSize();
      e.target.value = "";
    });
    addEventListener("resize", fit);
    // PowerPoint takes the pointer away from the add-in when a mouse button
    // goes down on it, which ends :hover before the click lands. Keep the bar
    // up (and clickable) for a moment after the pointer was last seen.
    let sleepTimer = 0;
    const wake = () => {
      clearTimeout(sleepTimer);
      document.body.classList.add("awake");
      sleepTimer = setTimeout(() => document.body.classList.remove("awake"), 2500);
    };
    for (const ev of ["pointermove", "pointerdown", "pointerup", "click"]) {
      document.addEventListener(ev, wake, true);
    }
    // The bar also shows while the box has focus (after a click into it).
    let focused = null;
    setInterval(() => {
      const now = document.hasFocus();
      if (now !== focused) document.body.classList.toggle("focused", (focused = now));
    }, 400);
    buildSamples();
  }

  /* ---------------- boot ---------------- */

  let started = false;

  function start(reason) {
    if (started) return;
    started = true;
    log("start", `${reason}; inOffice=${inOffice}`);
    try {
      wireUi();
      state = store.load();
      log("state", describe(state));
      render();
      log("rendered", `config ${config.hidden ? "hidden" : "shown"}, iframe ${stage.querySelector("iframe") ? "yes" : "no"}, ${innerWidth}x${innerHeight}`);
    } catch (e) {
      log("start-failed", e && (e.stack || e.message));
      throw e;
    }
  }

  function describe(s) {
    if (!s) return "empty -> setup panel";
    return `${s.mode} ${s.mode === "url" ? s.url : (s.name + ", " + s.html.length + " chars")} width=${s.width}`;
  }

  if (typeof Office !== "undefined" && Office.onReady) {
    log("office.js", "loaded, waiting for onReady");
    Office.onReady((info) => {
      inOffice = !!(info && info.host);
      log("onReady", `host=${info && info.host} platform=${info && info.platform}`);
      if (inOffice) watchView();
      if (!started) return start("onReady");
      // The fallback below already started from browser storage. Now that the
      // document is reachable, show what it holds instead.
      if (inOffice) {
        state = store.load();
        log("state", "late onReady: " + describe(state));
        if (state) closeConfig();
        render();
      }
    });
    // Never leave the box blank if Office is slow to initialise.
    setTimeout(() => { if (!started) { log("warn", "onReady not fired after 5s"); start("timeout"); } }, 5000);
  } else {
    log("office.js", "not loaded");
    start("no office.js"); // offline browser preview
  }
})();
