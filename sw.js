// バージョンを上げると古いキャッシュが自動削除される
const CACHE_NAME = 'katsudo-nippo-v7';
const CORE_ASSETS = [
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

// インストール：コアファイルのみキャッシュ（外部CDNは除外）
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(CORE_ASSETS))
      .then(() => self.skipWaiting()) // 常に即座に有効化
  );
});

// メッセージ受信（ページからの強制更新指示）
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

// アクティベート：古いバージョンのキャッシュを全削除
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => {
          console.log('[SW] Deleting old cache:', k);
          return caches.delete(k);
        })
      ))
      .then(() => self.clients.claim())
  );
});

// フェッチ戦略：
// - コアファイル → キャッシュ優先（オフライン対応）
// - CDN・外部リソース → ネットワーク優先（キャッシュはfallback）
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  const url = event.request.url;
  if (!url.startsWith('http')) return;

  const isCore = CORE_ASSETS.some(a => url.endsWith(a.replace('./','')));

  if (isCore) {
    // キャッシュ優先
    event.respondWith(
      caches.match(event.request).then(cached => {
        if (cached) return cached;
        return fetch(event.request).then(res => {
          const clone = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
          return res;
        });
      })
    );
  } else {
    // ネットワーク優先（fonts, cdnjs など）
    event.respondWith(
      fetch(event.request).catch(() => caches.match(event.request))
    );
  }
});
