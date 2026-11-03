// Service worker del Rapportino di Cantiere: conserva i file dell'app per aprirla anche senza campo.
// - File dell'app: "prima la rete" (con internet si prende sempre la versione aggiornata, altrimenti la copia).
// - Librerie di Firebase (gstatic.com): "prima la copia", perché ogni versione non cambia mai.
// - I dati (Firestore) non passano di qui: li gestisce Firebase con la sua cache.
// Il numero di versione sta in version.js: cambiandolo, il browser riconosce l'aggiornamento.
importScripts('version.js');
const CACHE = 'rapportino-' + self.APP_VERSION;
const CACHE_LIBRERIE = 'rapportino-librerie';
const FILES = [
  './', 'index.html', 'style.css', 'version.js', 'app.js', 'firebase-config.js', 'excel.js', 'cartella.js',
  'manifest.webmanifest', 'libs/jszip.min.js', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE && k !== CACHE_LIBRERIE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  if (url.origin === 'https://www.gstatic.com' && url.pathname.startsWith('/firebasejs/')) {
    event.respondWith(
      caches.open(CACHE_LIBRERIE).then(cache => cache.match(req).then(hit => hit || fetch(req).then(res => {
        if (res.ok) cache.put(req, res.clone());
        return res;
      })))
    );
    return;
  }

  if (url.origin !== self.location.origin) return;
  // "no-cache": chiede sempre al server se il file è cambiato (GitHub Pages altrimenti fa usare
  // per 10 minuti la copia già scaricata, e il telefono resta sulla versione vecchia)
  const fresca = req.mode === 'navigate'
    ? new Request(req.url, { cache: 'no-cache', credentials: 'same-origin' })
    : new Request(req, { cache: 'no-cache' });
  event.respondWith(
    fetch(fresca)
      .then(res => {
        const copy = res.clone();
        caches.open(CACHE).then(cache => cache.put(req, copy));
        return res;
      })
      .catch(() => caches.match(req, { ignoreSearch: true }))
  );
});
