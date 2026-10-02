// Offline cache for the app. Bump VERSION when any app file changes.
const VERSION = "ejk-v3";
const FILES = ["./", "./index.html", "./app.js", "./data.js", "./data.json", "./xlsx.full.min.js", "./manifest.webmanifest",
  "./icons/icon-192.png", "./icons/icon-512.png", "./icons/maskable-512.png", "./icons/apple-touch-icon.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(VERSION).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  const key = url.origin + url.pathname; // one cache entry per file, whatever the query string
  // network first so new draws and app updates arrive; cached copy when offline
  e.respondWith(fetch(e.request).then(r => {
    if (r.ok){ const copy = r.clone(); caches.open(VERSION).then(c => c.put(key, copy)); }
    return r;
  }).catch(() => caches.match(key).then(r => r || caches.match("./index.html"))));
});
