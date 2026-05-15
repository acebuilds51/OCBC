// OCBC Golf League Service Worker
// Version bumped with each app release — forces cache refresh
const CACHE_VERSION = 'ocbc-v1.59';
const CACHE_NAME = CACHE_VERSION;

// On install — cache the app shell
self.addEventListener('install', event => {
  // Skip waiting immediately so new SW takes over right away
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(['/OCBC/', '/OCBC/index.html', '/OCBC/manifest.json']);
    }).catch(() => {})
  );
});

// On activate — delete old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// On fetch — network first, fall back to cache
self.addEventListener('fetch', event => {
  // Always go to network for the HTML file so version updates are instant
  if(event.request.url.includes('index.html') || event.request.url.endsWith('/OCBC/') || event.request.url.endsWith('/OCBC')){
    event.respondWith(
      fetch(event.request).then(response => {
        const clone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        return response;
      }).catch(() => caches.match(event.request))
    );
    return;
  }
  // For everything else — cache first
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request))
  );
});

// Listen for skipWaiting message
self.addEventListener('message', event => {
  if(event.data && event.data.action === 'skipWaiting') self.skipWaiting();
});
