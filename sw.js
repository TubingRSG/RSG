// RSG — garde l'application sur l'appareil pour qu'elle s'ouvre même sans Internet.
// Les données, elles, viennent de la base (ou de la dernière copie gardée par l'application).
const CACHE = "rsg-v1";
const SHELL = ["./", "./index.html", "./logo.png", "./favicon.ico", "./icon-192.png", "./icon-512.png", "./manifest.json"];
const LIBS = [
  "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.49.4/dist/umd/supabase.js",
  "https://cdnjs.cloudflare.com/ajax/libs/qrcode-generator/1.4.4/qrcode.min.js",
];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL).then(() => Promise.all(LIBS.map(u => c.add(u).catch(() => {}))))));
  self.skipWaiting();
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  const mine = url.origin === location.origin;
  if (!mine && !LIBS.includes(req.url)) return;          // la base Supabase n'est jamais mise en cache ici
  if (req.mode === "navigate" || url.pathname.endsWith("/index.html")) {
    // Page principale : la version en ligne d'abord (mises à jour), sinon la copie gardée.
    e.respondWith(fetch(req).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put("./index.html", c)); return r; })
      .catch(() => caches.match("./index.html")));
    return;
  }
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => {
    if (r.ok) { const c = r.clone(); caches.open(CACHE).then(x => x.put(req, c)); }
    return r;
  })));
});
