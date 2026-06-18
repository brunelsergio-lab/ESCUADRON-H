const CACHE_NAME = 'escuadron-h-shell-v1';
const STATIC_ASSETS = [
  '/app/static/manifest.json',
  '/app/static/icon-192.png',
  '/app/static/icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS)).catch(() => null));
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  if (url.pathname.startsWith('/app/static/')) {
    event.respondWith(caches.match(event.request).then(resp => resp || fetch(event.request)));
  }
});
