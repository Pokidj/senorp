const CACHE_NAME = "senorp-pwa-v35";
const APP_ASSETS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./senorp_config.js",
  "./jspdf.umd.min.js",
  "./logo-mini.png",
  "./logo-transparent.png",
  "./logo-dark.png",
  "./pdf-logo.png"
];

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => Promise.all(APP_ASSETS.map(async asset => {
        try {
          const response = await fetch(asset, { cache: "reload" });
          if (response.ok) await cache.put(asset, response);
        } catch {
          // One optional asset must not prevent the PWA from installing.
        }
      })))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;
  const requestUrl = new URL(event.request.url);
  if (requestUrl.origin !== self.location.origin) return;

  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        if (response.ok) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        }
        return response;
      }).catch(() => event.request.mode === "navigate"
        ? caches.match("./index.html")
        : Response.error());
    })
  );
});
