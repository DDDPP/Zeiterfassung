// ── GARTEN HOCH 3 – Service Worker ────────────────────────────────────────────
// Version hier hochzählen bei jedem Update → löst automatisches Update aus
const CACHE_VERSION = 'gh3-v5';
const CACHE_FILES = [
  './zeiterfassung.html',
  './apple-touch-icon.png',
  './manifest.json'
];

// Installation: Dateien cachen
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then(cache => cache.addAll(CACHE_FILES))
  );
});

// Aktivierung: Alten Cache löschen
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k)))
    )
  );
});

// Fetch: Cache-first, dann Netzwerk
self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request))
  );
});

// Update auf Anfrage der App
self.addEventListener('message', event => {
  if (event.data && event.data.action === 'skipWaiting') {
    self.skipWaiting();
  }
});
