const CACHE_VERSION = 'v2';
const STATIC_CACHE = `static-${CACHE_VERSION}`;

const PRECACHE_URLS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './lib/exceljs.min.js',
  './lib/FileSaver.min.js',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

// Install: 必要ファイルを事前キャッシュ
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(STATIC_CACHE).then((cache) => cache.addAll(PRECACHE_URLS))
  );
  self.skipWaiting();
});

// Activate: 古いキャッシュの削除
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== STATIC_CACHE)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

// Fetch: ナビゲーションはネット優先→失敗時にindex.html、静的アセットはキャッシュ優先
self.addEventListener('fetch', (event) => {
  const req = event.request;

  // ページ遷移（HTML）はオンライン優先、失敗時はオフライン用にindex.html
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req).catch(() => caches.match('./index.html'))
    );
    return;
  }

  // 静的ファイルはキャッシュ優先、なければ取得してキャッシュへ
  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;
      return fetch(req)
        .then((res) => {
          // GET & 200系のみキャッシュ
          if (req.method === 'GET' && res && (res.status === 200 || res.type === 'opaque')) {
            const copy = res.clone();
            caches.open(STATIC_CACHE).then((cache) => cache.put(req, copy));
          }
          return res;
        })
        .catch(() => {
          // オフライン時のフォールバック（必要なら画像等に代替を返す）
          return caches.match(req);
        });
    })
  );
});