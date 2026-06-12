// 筒賀水質管理センター 週報 Service Worker
// v1: 筒賀アプリ 初期キャッシュ
const CACHE_NAME = 'tutuga-v36';
const APP_VERSION = '20260612-1';
const ASSETS = [
  './',
  './index.html?v=' + APP_VERSION,
  './excel-write-map.js?v=' + APP_VERSION,
  './template.js?v=' + APP_VERSION,
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS).catch(() => {}))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(
      keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
    )).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  // GAS API は常にネットワーク
  if (req.url.includes('script.google.com') || req.url.includes('googleusercontent.com')) {
    return;
  }
  // 同一オリジンの静的ファイルは「ネットワークファースト」
  if (req.method === 'GET' && new URL(req.url).origin === self.location.origin) {
    event.respondWith(
      fetch(req).then((res) => {
        if (res && res.status === 200 && res.type === 'basic') {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
        }
        return res;
      }).catch(() => caches.match(req))
    );
  }
});
