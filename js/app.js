// Versioned cache name (change number when updating the app)
const CACHE_NAME = "inventory-mobile-offline-v2";

// Files to cache for offline use
const urlsToCache = [
  "./",
  "./index.html",
  "./manifest.json",
  "./sample-data.json",
  "./css/style.css",
  "./js/app.js",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"
];

// Install event (cache files)
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log("Caching app files...");
      return cache.addAll(urlsToCache);
    })
  );
});

// Activate event (remove old caches)
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            console.log("Deleting old cache:", key);
            return caches.delete(key);
          }
        })
      )
    )
  );
});

// Fetch event (serve from cache if offline)
self.addEventListener("fetch", (event) => {
  event.respondWith(
    caches.match(event.request).then((response) => {
      return response || fetch(event.request);
    })
  );
});
