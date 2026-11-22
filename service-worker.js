
const CACHE_NAME = "portal-bmn-static-v1";

// Hanya menyimpan aset statis yang aman.
const STATIC_ASSETS = [
  "./manifest.json",
  "./logo-bmn.png.png",
  "./icons/icon-192.png",
  "./icons/icon-512.png"
];

// Memasang service worker.
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(STATIC_ASSETS))
      .then(() => self.skipWaiting())
  );
});

// Mengaktifkan service worker baru.
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key.startsWith("portal-bmn-static-") && key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

// Tidak melakukan cache terhadap halaman, login, API, atau data BMN.
self.addEventListener("fetch", (event) => {
  const request = event.request;

  if (request.method !== "GET") return;

  const url = new URL(request.url);

  if (url.origin !== self.location.origin) return;

  const isStaticAsset =
    url.pathname.endsWith(".png") ||
    url.pathname.endsWith(".jpg") ||
    url.pathname.endsWith(".jpeg") ||
    url.pathname.endsWith(".webp") ||
    url.pathname.endsWith(".svg");

  if (!isStaticAsset) return;

  event.respondWith(
    caches.match(request).then((cached) =>
      cached || fetch(request)
    )
  );
});
