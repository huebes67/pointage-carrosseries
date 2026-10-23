// Garde l'interface en cache pour un démarrage instantané (les données viennent toujours d'Excel).
// Changer VERSION à chaque mise en ligne pour forcer la mise à jour sur les téléphones.
const VERSION = "pc-v2";
const SHELL = [
  "./", "index.html", "style.css", "app.js", "config.js", "manifest.webmanifest",
  "vendor/xlsx.full.min.js",
  "icons/icon-192.png", "icons/apple-touch-icon.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// Réseau d'abord (pour avoir la dernière version), cache si hors ligne.
// Seuls les fichiers de l'appli sont concernés : Microsoft / Graph ne passent jamais par le cache.
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  e.respondWith(
    fetch(e.request)
      .then((res) => {
        const copy = res.clone();
        caches.open(VERSION).then((c) => c.put(e.request, copy));
        return res;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
