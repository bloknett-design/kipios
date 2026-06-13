// ============================================================
// Service Worker для КИПиА — стратегия network-first
// ============================================================
// КАК ОБНОВИТЬ САЙТ: просто измените CACHE_VERSION ниже.
// При каждом изменении файлов на сервере увеличивайте версию:
//   'kipia-v1' → 'kipia-v2' → 'kipia-v3' и т.д.
// Браузер увидит новый SW → удалит старый кэш → загрузит
// свежие файлы с сервера.
// ============================================================

const CACHE_VERSION = 'kipia-v12';
const CACHE_NAME = CACHE_VERSION;

const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon.png'
];

// Install — предварительное кэширование основных файлов
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  // Немедленно активировать новый SW, не дожидаясь закрытия старых вкладок
  self.skipWaiting();
});

// Activate — удалить старые кэши
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  // Захватить контроль над всеми вкладками сразу
  self.clients.claim();
});

// Fetch — NETWORK-FIRST для локальных файлов, network-first для внешних
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Только GET-запросы
  if (event.request.method !== 'GET') return;

  // Внешние ресурсы (шрифты, CDN) — network-first, fallback to cache
  if (url.origin !== self.location.origin) {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          }
          return response;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // Локальные файлы — NETWORK-FIRST (сначала сервер, при ошибке — кэш)
  event.respondWith(
    fetch(event.request)
      .then(response => {
        if (response.ok) {
          // Обновляем кэш свежим ответом
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => {
        // Нет сети — отдаём из кэша
        return caches.match(event.request).then(cached => {
          return cached || new Response('Offline', { status: 503, statusText: 'Service Unavailable' });
        });
      })
  );
});
