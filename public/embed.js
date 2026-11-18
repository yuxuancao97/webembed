/* WebEmbed: on-slide logic for the PowerPoint content add-in.
 *
 * Each inserted add-in is its own instance with its own settings, so one deck
 * can hold many embeds. What gets stored (in the .pptx) is:
 *   { mode: "url",  url,  zoom }    a live page, fetched every time the slide shows
 *   { mode: "html", html, name, zoom }  inline HTML, rendered in a sandboxed iframe
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

  const $ = (sel) => document.querySelector(sel);
  const stage = $("#stage");
  const config = $("#config");
  const toolbar = $("#toolbar");
  const urlInput = $("#url-input");
  const urlStatus = $("#url-status");
  const htmlInput = $("#html-input");
  const htmlSize = $("#html-size");
  const zoomSel = $("#zoom");
  const errorEl = $("#error");

  let inOffice = false;
  let state = null;
  let activeTab = "url";

  /* ---------------- storage ---------------- */

  const store = {
    load() {
      if (inOffice) return Office.context.document.settings.get(KEY) || null;
      try { return JSON.parse(localStorage.getItem(KEY)); } catch { return null; }
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
    (window.__wlog || (() => {}))("view", view);
  }

  function watchView() {
    const doc = Office.context.document;
    if (typeof doc.getActiveViewAsync === "function") {
      doc.getActiveViewAsync((r) => { if (r.status === "succeeded") setView(r.value); });
    }
    if (Office.EventType && Office.EventType.ActiveViewChanged && doc.addHandlerAsync) {
      try { doc.addHandlerAsync(Office.EventType.ActiveViewChanged, (e) => setView(e.activeView)); }
      catch { /* not supported: toolbar just stays hover-only */ }
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

  // Asks the add-in server whether the site allows framing (X-Frame-Options / CSP).
  // The browser can't tell us this itself: a blocked frame still fires "load".
  async function checkFraming(url) {
    try {
      const r = await fetch("api/check?url=" + encodeURIComponent(url) + "&origin=" + encodeURIComponent(location.origin), { cache: "no-store" });
      if (!r.ok) return null;
      return await r.json();
    } catch {
      return null; // static hosting without the check endpoint
    }
  }

  /* ---------------- rendering ---------------- */

  function render() {
    stage.replaceChildren();
    document.body.classList.toggle("has-content", !!state);
    toolbar.hidden = !state;
    if (!state) return openConfig();

    const f = document.createElement("iframe");
    if (state.mode === "url") {
      f.src = state.url;
      f.allow = "autoplay; fullscreen; encrypted-media; picture-in-picture; clipboard-write";
      f.referrerPolicy = "strict-origin-when-cross-origin";
      f.title = state.url;
      f.addEventListener("load", () => log("iframe-load", state && state.url));
    } else {
      // No allow-same-origin: inline HTML runs with an opaque origin and
      // cannot reach the add-in or Office APIs.
      f.setAttribute("sandbox", "allow-scripts allow-popups");
      f.allow = "autoplay; fullscreen";
      f.srcdoc = state.html;
      f.title = state.name || "Embedded HTML";
    }
    const z = Number(state.zoom) || 1;
    if (z !== 1) {
      f.style.width = (100 / z) + "%";
      f.style.height = (100 / z) + "%";
      f.style.transform = `scale(${z})`;
    }
    stage.append(f);

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

  function showBlocked(url, reason) {
    let host = url;
    try { host = new URL(url).hostname.replace(/^www\./, ""); } catch {}
    const card = document.createElement("div");
    card.className = "blocked";
    card.innerHTML = `
      <strong></strong>
      <p class="why"></p>
      <div class="actions">
        <button type="button" class="primary" data-act="open">Open in browser</button>
        <button type="button" class="quiet edit-only" data-act="change">Change</button>
      </div>`;
    card.querySelector("strong").textContent = `${host} can't be shown inside a slide`;
    card.querySelector(".why").textContent = reason + " Only the site's owner can allow it.";
    card.querySelector('[data-act="open"]').onclick = () => openInBrowser(url);
    card.querySelector('[data-act="change"]').onclick = openConfig;
    stage.replaceChildren(card);
    document.body.classList.remove("has-content");
    log("blocked", `${url}: ${reason}`);
  }

  /* ---------------- config panel ---------------- */

  function selectTab(tab) {
    activeTab = tab;
    for (const b of config.querySelectorAll("[role=tab]")) b.setAttribute("aria-selected", String(b.dataset.tab === tab));
    for (const p of config.querySelectorAll(".pane")) p.hidden = p.dataset.pane !== tab;
    $("#btn-apply").hidden = tab === "samples";
    showError("");
  }

  function openConfig() {
    const has = !!state;
    $("#btn-cancel").hidden = !has;
    $("#btn-remove").hidden = !has;
    zoomSel.value = String(state?.zoom ?? 1);
    if (state?.mode === "url") urlInput.value = state.url;
    if (state?.mode === "html") htmlInput.value = state.html;
    updateSize();
    selectTab(state?.mode || activeTab);
    config.hidden = false;
    toolbar.hidden = true;
  }

  function closeConfig() {
    config.hidden = true;
    toolbar.hidden = !state;
  }

  function showError(msg) {
    errorEl.textContent = msg;
    errorEl.hidden = !msg;
  }

  // Measure what is actually stored: the JSON-escaped string.
  const byteLength = (s) => new TextEncoder().encode(JSON.stringify(s)).length;

  function updateSize() {
    const n = byteLength(htmlInput.value);
    if (!n) { htmlSize.textContent = ""; return; }
    const kb = (n / 1024).toFixed(n < 10240 ? 1 : 0);
    const over = n > MAX_INLINE_BYTES;
    htmlSize.textContent = over
      ? `${kb} KB: too big to store in the deck (limit ${Math.round(MAX_INLINE_BYTES / 1024)} KB). Host it and embed by URL instead.`
      : `${kb} KB`;
    htmlSize.className = "hint" + (over ? " warn" : "");
  }

  async function commit(next) {
    const log = window.__wlog || (() => {});
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

  async function applyUrl() {
    let url;
    try { url = normalizeUrl(urlInput.value); } catch (e) { return showError(e.message); }
    urlInput.value = url;
    urlStatus.className = "hint";
    urlStatus.textContent = "Checking whether the site allows embedding…";
    const verdict = await checkFraming(url);
    if (verdict && verdict.embeddable === false) {
      urlStatus.className = "hint warn";
      urlStatus.textContent = `${verdict.reason} It will most likely show as blank or an error. Press Embed again to use it anyway.`;
      if (urlInput.dataset.warned !== url) { urlInput.dataset.warned = url; return; }
    } else if (verdict && verdict.embeddable === true) {
      urlStatus.className = "hint ok";
      urlStatus.textContent = "Looks embeddable.";
    } else {
      urlStatus.className = "hint";
      urlStatus.textContent = verdict?.reason || "Couldn't check this site. If the slide stays blank, the site doesn't allow embedding.";
    }
    await commit({ mode: "url", url, zoom: Number(zoomSel.value) });
  }

  async function applyHtml(html, name) {
    if (!html.trim()) return showError("Paste some HTML or open a file first.");
    if (byteLength(html) > MAX_INLINE_BYTES) {
      return showError("Too big to store in the deck. Host the file and embed it by URL instead.");
    }
    await commit({ mode: "html", html, name: name || "Inline HTML", zoom: Number(zoomSel.value) });
  }

  function buildSamples() {
    const list = $("#sample-list");
    for (const s of SAMPLES) {
      const li = document.createElement("li");
      li.innerHTML = `<span class="name"></span><button type="button" data-act="link">Link</button><button type="button" data-act="copy">Copy into deck</button>`;
      li.querySelector(".name").innerHTML = `${s.title}<small>${s.note}</small>`;
      li.querySelector('[data-act="link"]').onclick = () =>
        commit({ mode: "url", url: new URL(s.file, location.href).href, zoom: Number(zoomSel.value) });
      li.querySelector('[data-act="copy"]').onclick = async () => {
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
    $("#btn-apply").onclick = () => (activeTab === "url" ? applyUrl() : applyHtml(htmlInput.value));
    $("#btn-cancel").onclick = () => { showError(""); closeConfig(); };
    $("#btn-remove").onclick = () => { urlInput.value = ""; htmlInput.value = ""; commit(null); };
    $("#btn-edit").onclick = openConfig;
    $("#btn-reload").onclick = render;
    urlInput.addEventListener("keydown", (e) => { if (e.key === "Enter") applyUrl(); });
    urlInput.addEventListener("input", () => { delete urlInput.dataset.warned; });
    htmlInput.addEventListener("input", updateSize);
    $("#html-file").addEventListener("change", async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      htmlInput.value = await file.text();
      updateSize();
      e.target.value = "";
    });
    buildSamples();
  }

  /* ---------------- boot ---------------- */

  const log = window.__wlog || (() => {});
  let started = false;

  function start(reason) {
    if (started) return;
    started = true;
    log("start", `${reason}; inOffice=${inOffice}`);
    try {
      wireUi();
      state = store.load();
      log("state", state ? `${state.mode} ${state.mode === "url" ? state.url : (state.name + ", " + state.html.length + " chars")}` : "empty -> setup panel");
      render();
      log("rendered", `config ${config.hidden ? "hidden" : "shown"}, iframe ${stage.querySelector("iframe") ? "yes" : "no"}, ${innerWidth}x${innerHeight}`);
    } catch (e) {
      log("start-failed", e && (e.stack || e.message));
      throw e;
    }
  }

  if (typeof Office !== "undefined" && Office.onReady) {
    log("office.js", "loaded, waiting for onReady");
    Office.onReady((info) => {
      inOffice = !!(info && info.host);
      log("onReady", `host=${info && info.host} platform=${info && info.platform}`);
      if (inOffice) watchView();
      start("onReady");
    });
    // Never leave the box blank if Office is slow to initialise.
    setTimeout(() => { if (!started) { log("warn", "onReady not fired after 5s"); start("timeout"); } }, 5000);
  } else {
    log("office.js", "not loaded");
    start("no office.js"); // offline browser preview
  }
})();
