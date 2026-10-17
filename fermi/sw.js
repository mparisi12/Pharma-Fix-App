// Service worker: l'app si apre anche senza rete (i dati restano in coda e partono appena c'è campo).
// Rete prima, cache come riserva: così ogni aggiornamento pubblicato su GitHub arriva subito.
const CACHE = 'pf-v5';
const SHELL = ['./', './index.html', './config.js?v=20261017', './defaults.js', './manifest.webmanifest', './icon-192.png', './apple-touch-icon.png',
  'https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(() => {})); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))); self.clients.claim(); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.hostname.includes('script.google')) return; // mai in cache le chiamate al foglio
  // file dell'app: chiedi sempre a GitHub se sono cambiati (niente versioni vecchie dalla cache del browser)
  const req = u.origin === location.origin ? fetch(e.request.url, { cache: 'no-cache', credentials: 'same-origin' }) : fetch(e.request);
  e.respondWith(req.then(r => {
    if (r.ok && (u.origin === location.origin || u.hostname === 'cdn.jsdelivr.net')) { const cp = r.clone(); caches.open(CACHE).then(c => c.put(e.request, cp)); }
    return r;
  }).catch(() => caches.match(e.request, { ignoreSearch: true })));
});
