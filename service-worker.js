const CACHE_NAME = 'stockroom-tracker-v1';
const SHELL_FILES = [
  './index.html',
  './manifest.json'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

// Network-first for everything (this app is live data, not content to cache),
// falling back to the cached app shell only if the device is offline.
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return; // never cache POST calls to the Sheet API
  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  );
});
