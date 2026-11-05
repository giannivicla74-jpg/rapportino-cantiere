// Service worker del Rapportino di Cantiere: conserva i file dell'app per aprirla anche senza campo.
// - File dell'app: "prima la copia", così con il campo debole l'app si apre subito e non resta ferma sull'icona.
//   La rete serve solo a scaricare in sottofondo una versione nuova, tutta insieme.
// - Librerie di Firebase (gstatic.com): "prima la copia", perché ogni versione non cambia mai.
// - I dati (Firestore) non passano di qui: li gestisce Firebase con la sua cache.
// Il numero di versione sta in version.js: cambiandolo, il browser riconosce l'aggiornamento.
importScripts('version.js');
const CACHE = 'rapportino-' + self.APP_VERSION;
const CACHE_LIBRERIE = 'rapportino-librerie';
// Le icone grandi le scarica il telefono per conto suo: qui non servono e renderebbero pesante l'aggiornamento.
const FILES = [
  './', 'index.html', 'style.css', 'version.js', 'app.js', 'firebase-config.js', 'excel.js', 'cartella.js', 'archivio.js',
  'manifest.webmanifest', 'libs/jszip.min.js', 'icons/icon-192.png', 'icons/logo-vicla.png', 'icons/gc-codelab.png'
];

self.addEventListener('install', event => {
  // "reload": file presi davvero dal server, mai da una copia vecchia del browser
  event.waitUntil(caches.open(CACHE)
    .then(cache => cache.addAll(FILES.map(f => new Request(f, { cache: 'reload' }))))
    .then(() => self.skipWaiting()));
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
  // la rete si usa solo per ciò che non è in copia (per esempio le icone grandi)
  event.respondWith(
    caches.open(CACHE).then(cache => cache.match(req, { ignoreSearch: true }).then(hit => hit || fetch(req)))
  );
});
