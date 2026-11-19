/* Service worker Simarisa-sakti. Naikkan nomor VERSI jika ingin memaksa cache lama dibuang. */
const VERSI = 'simarisa-v2';
const SHELL = ['/', '/favicon.svg', '/icon-192.png', '/icon-512.png', '/apple-touch-icon.png', '/site.webmanifest'];
const CDN = ['cdn.jsdelivr.net', 'fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSI).then(c => Promise.all(SHELL.map(u => c.add(u).catch(() => {})))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSI).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

// Data aplikasi (Apps Script, Google Drive) tidak pernah disimpan di cache: hanya cangkang aplikasi dan pustaka CDN.
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (r.mode === 'navigate' && u.origin === location.origin) { // halaman: utamakan versi terbaru, cadangan dari cache saat offline
    e.respondWith(fetch(r).then(res => { if (res.ok) { const cp = res.clone(); caches.open(VERSI).then(c => c.put('/', cp)); } return res; })
      .catch(() => caches.match('/').then(m => m || new Response('Anda sedang offline.', { status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8' } }))));
    return;
  }
  if (u.origin === location.origin || CDN.indexOf(u.hostname) > -1) { // aset statis: tampilkan cache dulu, perbarui di latar belakang
    e.respondWith(caches.open(VERSI).then(c => c.match(r).then(m => {
      const net = fetch(r).then(res => { if (res.ok || res.type === 'opaque') c.put(r, res.clone()); return res; }).catch(() => m);
      return m || net;
    })));
  }
});
