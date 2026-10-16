/* Platform365 service worker — Present + EasyWrite, sürüm 2.0.0 */
const VERSION = '2.0.0';
const BUILD = 1; // Dosyalar değiştiğinde artırın; eski önbellek silinir
const CACHE = `platform365-v${VERSION}-${BUILD}`;
const ASSETS = [
  './',
  'index.html',
  'hub.css',
  'hub.js',
  'manifest.webmanifest',
  'shared/platform.css',
  'shared/platform.js',
  'shared/convert.js',
  'present/',
  'present/index.html',
  'present/styles.css',
  'present/app.js',
  'present/features.js',
  'present/shapes.js',
  'present/pptx-import.js',
  'present/pptx-export.js',
  'easywrite/',
  'easywrite/index.html',
  'easywrite/styles.css',
  'easywrite/app.js',
  'easywrite/docx.js',
  'vendor/pptxgen.bundle.js',
  'icons/icon.svg',
  'icons/easywrite.svg',
  'icons/platform365.svg',
  'icons/icon-maskable.svg',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/icon-maskable-512.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE)
      .then((c) => c.addAll(ASSETS.map((u) => new Request(u, { cache: 'reload' }))))
      .then(() => self.skipWaiting()),
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => (k.startsWith('present-') || k.startsWith('platform365-')) && k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

// Sayfa isteğinin hangi uygulamaya ait olduğunu bul (çevrimdışı yedek için)
const shellFor = (url) => {
  const p = new URL(url).pathname;
  if (/\/present\/?/.test(p)) return 'present/index.html';
  if (/\/easywrite\/?/.test(p)) return 'easywrite/index.html';
  return 'index.html';
};

// Önce ağ (her zaman güncel kod), çevrimdışıyken önbellek
self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  const nav = req.mode === 'navigate';
  event.respondWith(
    fetch(req, { cache: 'no-cache' })
      .then((res) => {
        if (res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(nav ? shellFor(req.url) : req, copy));
        }
        return res;
      })
      .catch(() => caches.match(nav ? shellFor(req.url) : req, { ignoreSearch: true })
        .then((hit) => hit || (nav ? caches.match('index.html') : undefined))),
  );
});
