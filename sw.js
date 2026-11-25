// Offline-ondersteuning: eigen bestanden netwerk-eerst (zodat updates meteen binnenkomen), bibliotheken cache-eerst.
const CACHE = 'vdp-uren-v2';
const SHELL = ['./', 'index.html', 'styles.css', 'app.js', 'config.js', 'logo.png', 'manifest.webmanifest',
  'icon-180.png', 'icon-192.png', 'icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (url.hostname.endsWith('microsoft.com') || url.hostname.endsWith('microsoftonline.com') ||
      url.hostname.endsWith('sharepoint.com') || url.hostname.endsWith('live.com')) return;   // nooit Graph/aanmelding cachen
  const isLib = url.hostname === 'cdn.jsdelivr.net' || url.hostname === 'cdnjs.cloudflare.com';
  if (isLib) {
    e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request).then(res => {
      const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return res;
    })));
  } else if (url.origin === self.location.origin) {
    e.respondWith(fetch(e.request).then(res => {
      const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return res;
    }).catch(() => caches.match(e.request, { ignoreSearch: true })));
  }
});
