// service-worker.js (root)
const CACHE_VERSION = 'v9';
const STATIC_CACHE = `static-${CACHE_VERSION}`;

// サイトルート基準のパスを列挙（存在しないものは取り除いてOK）
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
  './FieldSurveyApp/assets/js/modules/idb-photos.js', // IndexedDB ヘルパー（導入済みなら）
  './FieldSurveyApp/icons/icon-192.png',
  './FieldSurveyApp/icons/icon-512.png',
];

// 404をスキップする安全なプリキャッシュ
async function safePrecache() {
  const cache = await caches.open(STATIC_CACHE);
  const base = self.registration.scope; // 例: https://<user>.github.io/field-survey-app/
  const urls = PRECACHE_PATHS.map((p) => new URL(p, base).href);

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

// Install: プリキャッシュ（skipWaitingはここではしない→ユーザー確認で適用）
self.addEventListener('install', (event) => {
  event.waitUntil(safePrecache());
});

// Activate: 古いキャッシュ削除＋即時制御開始
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== STATIC_CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch: ナビゲーションはネット優先→失敗時index.html、静的はキャッシュ優先
self.addEventListener('fetch', (event) => {
  const req = event.request;

  // ページ遷移（HTML）
  if (req.mode === 'navigate') {
    event.respondWith(
      (async () => {
        try {
          const res = await fetch(req);
          // 404等でもSPA/PWAとして index.html に戻す
          if (res && res.ok) return res;
        } catch (e) {
          // ネットワーク失敗時もフォールバック
        }
        return caches.match(new URL('./index.html', self.registration.scope).href);
      })()
    );
    return;
  }

  // 静的ファイル
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

// メッセージ（UIからのSKIP_WAITING指示を受けて即時更新）
self.addEventListener('message', (e) => {
  if (e.data && e.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
