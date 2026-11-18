// WebEmbed dev server: serves ./public and a small /api/check endpoint.
// No dependencies. Uses HTTPS when the Office dev certs are installed
// (npm run certs), otherwise plain HTTP, which is fine for browser preview
// but not for PowerPoint.
import http from "node:http";
import https from "node:https";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const PORT = Number(process.env.PORT) || 3000;
const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), "public");
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");
const CERT = path.join(CERT_DIR, "localhost.crt");
const KEY = path.join(CERT_DIR, "localhost.key");
const FORCE_HTTP = process.argv.includes("--http");

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".webp": "image/webp",
  ".mp4": "video/mp4",
  ".webm": "video/webm",
  ".woff2": "font/woff2",
  ".json5": "application/json",
};

function send(res, status, body, type = "text/plain; charset=utf-8") {
  res.writeHead(status, { "Content-Type": type, "Cache-Control": "no-store" });
  res.end(body);
}

// ---- framing check -------------------------------------------------------
// Fetches the page server-side and reads X-Frame-Options and the CSP
// frame-ancestors directive, which decide whether a browser will render it
// inside an iframe on another site.
// Does a CSP source expression (https://host, https://*.host, host:port) match origin?
function sourceMatches(src, origin) {
  if (!origin) return false;
  let o;
  try { o = new URL(origin); } catch { return false; }
  const m = src.match(/^(?:(https?):\/\/)?(\*\.)?([^/:]+)(?::(\d+|\*))?\/?$/);
  if (!m) return false;
  const [, scheme, wild, host, port] = m;
  if (scheme && scheme + ":" !== o.protocol) return false;
  if (wild ? !o.hostname.endsWith("." + host) : o.hostname !== host) return false;
  const dflt = (p) => (p === "http:" ? "80" : "443");
  const oport = o.port || dflt(o.protocol);
  if (port === "*") return true;
  return oport === (port || dflt(o.protocol));
}

function verdictFrom(headers, finalUrl, origin) {
  const xfo = (headers.get("x-frame-options") || "").toLowerCase().trim();
  const csp = headers.get("content-security-policy") || "";
  const fa = csp.split(";").map((d) => d.trim()).find((d) => /^frame-ancestors\b/i.test(d));

  // When frame-ancestors is present, browsers ignore X-Frame-Options.
  if (fa) {
    const sources = fa.split(/\s+/).slice(1).map((s) => s.toLowerCase());
    if (sources.includes("*") || sources.includes("https:")) {
      return { embeddable: true, reason: "The site allows embedding anywhere." };
    }
    if (sources.some((s) => sourceMatches(s, origin))) {
      return { embeddable: true, reason: `The site allows embedding in this add-in (${origin}).` };
    }
    if (sources.length === 0 || sources.includes("'none'")) {
      return { embeddable: false, reason: "This site forbids embedding (CSP frame-ancestors 'none')." };
    }
    return { embeddable: false, reason: `This site only allows embedding on: ${sources.join(" ")}.` };
  }
  if (xfo === "deny") return { embeddable: false, reason: "This site forbids embedding (X-Frame-Options: DENY)." };
  if (xfo === "sameorigin") return { embeddable: false, reason: "This site only allows embedding on its own pages (X-Frame-Options: SAMEORIGIN)." };
  return { embeddable: true, reason: "No embedding restrictions found.", finalUrl };
}

async function check(target, origin, res) {
  let u;
  try { u = new URL(target); } catch { return send(res, 400, JSON.stringify({ error: "bad url" }), TYPES[".json"]); }
  if (!/^https?:$/.test(u.protocol)) return send(res, 400, JSON.stringify({ error: "http(s) only" }), TYPES[".json"]);

  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), 8000);
  try {
    const r = await fetch(u, {
      redirect: "follow",
      signal: ctl.signal,
      headers: { "User-Agent": "Mozilla/5.0 (Macintosh) WebEmbed-check", Accept: "text/html,*/*" },
    });
    r.body?.cancel().catch(() => {});
    const v = verdictFrom(r.headers, r.url, origin);
    if (r.status >= 400) {
      v.embeddable = null;
      v.reason = `The site answered ${r.status}; it may still load in the slide.`;
    }
    send(res, 200, JSON.stringify({ status: r.status, ...v }), TYPES[".json"]);
  } catch (e) {
    send(res, 200, JSON.stringify({ embeddable: null, reason: "Couldn't reach the site from the add-in server." }), TYPES[".json"]);
  } finally {
    clearTimeout(timer);
  }
}

// ---- static files ----------------------------------------------------------
function serveStatic(pathname, res) {
  if (pathname === "/") pathname = "/embed.html";
  const file = path.normalize(path.join(ROOT, decodeURIComponent(pathname)));
  if (!file.startsWith(ROOT + path.sep)) return send(res, 403, "Forbidden");
  fs.stat(file, (err, st) => {
    if (err || !st.isFile()) return send(res, 404, "Not found");
    res.writeHead(200, {
      "Content-Type": TYPES[path.extname(file).toLowerCase()] || "application/octet-stream",
      "Content-Length": st.size,
      "Cache-Control": "no-cache",
    });
    fs.createReadStream(file).pipe(res);
  });
}

// Receives diagnostics from the add-in page (see the inline script in embed.html)
// so problems inside PowerPoint show up in this terminal.
function logFromClient(req, res) {
  let body = "";
  req.on("data", (c) => { if (body.length < 8000) body += c; });
  req.on("end", () => {
    let m;
    try { m = JSON.parse(body); } catch { m = { kind: "raw", msg: body }; }
    console.log(`[addin ${m.id || "?"}] ${m.kind}: ${m.msg}`);
    send(res, 204, "");
  });
}

function handler(req, res) {
  const url = new URL(req.url, "http://localhost");
  if (url.pathname === "/api/log" && req.method === "POST") return logFromClient(req, res);
  if (url.pathname === "/api/check") return check(url.searchParams.get("url") || "", url.searchParams.get("origin") || "", res);
  if (url.pathname.endsWith(".html")) console.log(`GET ${url.pathname}  (${(req.headers["user-agent"] || "").slice(0, 90)})`);
  if (req.method !== "GET" && req.method !== "HEAD") return send(res, 405, "Method not allowed");
  serveStatic(url.pathname, res);
}

const haveCerts = fs.existsSync(CERT) && fs.existsSync(KEY);
if (haveCerts && !FORCE_HTTP) {
  https.createServer({ cert: fs.readFileSync(CERT), key: fs.readFileSync(KEY) }, handler)
    .listen(PORT, () => console.log(`WebEmbed ready: https://localhost:${PORT}/embed.html`));
} else {
  http.createServer(handler).listen(PORT, () => {
    console.log(`WebEmbed (HTTP preview) at http://localhost:${PORT}/embed.html`);
    if (!haveCerts) console.log("PowerPoint needs HTTPS: run `npm run certs` once, then restart.");
  });
}
