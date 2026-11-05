/* Service worker: guarda la app en el teléfono para abrirla sin internet. */
const VERSION = "tareas-vo-v1.0.0";
const SHELL = ["./", "index.html", "app.css", "app.js", "manifest.webmanifest", "guia.html",
  "icons/icon-192.png", "icons/icon-512.png", "icons/apple-touch-icon.png", "icons/maskable-512.png"];
const FONTS = "https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Barlow:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap";

self.addEventListener("install", e => {
  e.waitUntil((async () => {
    const c = await caches.open(VERSION);
    await c.addAll(SHELL);
    try { const r = await fetch(FONTS, {mode:"no-cors"}); await c.put(FONTS, r); } catch(_) {}
    self.skipWaiting();
  })());
});
self.addEventListener("activate", e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== VERSION) await caches.delete(k);
    await self.clients.claim();
  })());
});
self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;                       // las llamadas al flujo (POST) nunca se guardan
  const url = new URL(req.url);
  const isFont = url.hostname === "fonts.googleapis.com" || url.hostname === "fonts.gstatic.com";
  if (url.origin !== location.origin && !isFont) return;
  e.respondWith((async () => {
    const c = await caches.open(VERSION);
    const cached = await c.match(req, {ignoreSearch: req.mode === "navigate"}) || (req.mode === "navigate" ? await c.match("index.html") : null);
    const network = fetch(req).then(r => { if (r && (r.ok || r.type === "opaque")) c.put(req, r.clone()); return r; }).catch(() => null);
    if (cached) { e.waitUntil(network); return cached; }      // primero lo guardado, se refresca en segundo plano
    const r = await network;
    return r || new Response("Sin conexión", {status: 503, headers: {"Content-Type": "text/plain; charset=utf-8"}});
  })());
});
