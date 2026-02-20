// service-worker.js (root)
// scope: https://<ユーザー名>.github.io/field-survey-app/
const CACHE_VERSION = 'v7';
const STATIC_CACHE = `static-${CACHE_VERSION}`;

const PRECACHE_PATHS = [
  './',
  './index.html',
  './FieldSurveyApp/manifest.webmanifest',
  './lib/exceljs.min.js',
  './lib/FileSaver.min.js',
  './FieldSurveyApp/assets/css/styles.css',
  './FieldSurveyApp/assets/js/main.js',
  './FieldSurveyApp/assets/js/app.js',
  './FieldSurveyApp/assets/js/modules/pwa.js',
  './FieldSurveyApp/assets/js/modules/state.js',
  './FieldSurveyApp/assets/js/modules/storage.js',
  './FieldSurveyApp/assets/js/modules/utils.js',
  './FieldSurveyApp/assets/js/modules/excel/export.js',
  './FieldSurveyApp/assets/js/modules/excel/import.js',
  './FieldSurveyApp/assets/js/modules/ui/blocks.js',
  './FieldSurveyApp/assets/js/modules/ui/categories.js',
  './FieldSurveyApp/assets/js/modules/ui/fields.js',
  './FieldSurveyApp/assets/js/modules/ui/schemas.js',
  './FieldSurveyApp/icons/icon-192.png',
  './FieldSurveyApp/icons/icon-512.png',
];

async function safePrecache() {
  const cache = await caches.open(STATIC_CACHE);
  const base = self.registration.scope;
  const urls = PRECACHE_PATHS.map(p => new URL(p, base).href);
  await Promise.all(
    urls.map(async (url) => {
      try {
        const res = await fetch(url, { cache: 'no-cache' });
        if (res.ok || res.type === 'opaque') {
          await cache.put(url, res.clone());
        } else {
          console.warn('Precache skip (bad response):', url, res.status);
        }
      } catch (e) {
        console.warn('Precache failed:', url, e);
      }
    })
  );
}

self.addEventListener('install', (event) => {
  event.waitUntil(safePrecache());
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== STATIC_CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req).catch(() => caches.match(new URL('./index.html', self.registration.scope).href))
    );
    return;
  }
  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;
      return fetch(req)
        .then((res) => {
          if (req.method === 'GET' && res && (res.status === 200 || res.type === 'opaque')) {
            caches.open(STATIC_CACHE).then((cache) => cache.put(req, res.clone()));
          }
          return res;
        })
        .catch(() => caches.match(req));
    })
  );
});
