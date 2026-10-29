const V = 'trip-v1';
const SHELL = ['./', 'index.html', 'config.js', 'manifest.webmanifest', 'icon-180.png', 'icon-192.png', 'icon-512.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // App files: try network first so updates show up, fall back to cache when offline
  if (url.origin === self.location.origin) {
    e.respondWith(fetch(req).then(res => {
      const copy = res.clone(); caches.open(V).then(c => c.put(req, copy)); return res;
    }).catch(() => caches.match(req).then(r => r || caches.match('index.html'))));
    return;
  }
  // Fonts: cache first
  if (/fonts\.(googleapis|gstatic)\.com$/.test(url.hostname)) {
    e.respondWith(caches.match(req).then(r => r || fetch(req).then(res => {
      const copy = res.clone(); caches.open(V).then(c => c.put(req, copy)); return res;
    })));
  }
  // Everything else (the Google Sheet API) goes straight to the network
});
