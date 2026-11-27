/* 煎心鍋貼 PWA Service Worker
 * - 頁面：網路優先，離線時改用快取
 * - 圖示等本站檔案：快取優先
 * - CDN（樣式、字型、圖示、程式庫、圖片）：先用快取、背景更新
 * 修改網站後若要強制更新快取，把 VERSION 加 1 即可。
 */
var VERSION = 'v1';
var SHELL = 'jianxin-shell-' + VERSION;
var RUNTIME = 'jianxin-runtime-' + VERSION;
var SHELL_FILES = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png'
];
var CDN_HOSTS = [
  'cdn.tailwindcss.com', 'cdn.jsdelivr.net', 'cdnjs.cloudflare.com',
  'fonts.googleapis.com', 'fonts.gstatic.com', 'images.unsplash.com', 'i.postimg.cc'
];

self.addEventListener('install', function (event) {
  event.waitUntil(caches.open(SHELL).then(function (c) { return c.addAll(SHELL_FILES); }).then(function () { return self.skipWaiting(); }));
});

self.addEventListener('activate', function (event) {
  event.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== SHELL && k !== RUNTIME; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

self.addEventListener('fetch', function (event) {
  var req = event.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);

  // 頁面導覽：網路優先
  if (req.mode === 'navigate') {
    event.respondWith(fetch(req).then(function (res) {
      var copy = res.clone();
      caches.open(SHELL).then(function (c) { c.put('./index.html', copy); });
      return res;
    }).catch(function () {
      return caches.match('./index.html');
    }));
    return;
  }

  // 本站檔案：快取優先
  if (url.origin === self.location.origin) {
    event.respondWith(caches.match(req).then(function (hit) {
      return hit || fetch(req).then(function (res) {
        if (res.ok) { var copy = res.clone(); caches.open(RUNTIME).then(function (c) { c.put(req, copy); }); }
        return res;
      });
    }));
    return;
  }

  // CDN：先用快取、背景更新
  if (CDN_HOSTS.indexOf(url.hostname) !== -1) {
    event.respondWith(caches.open(RUNTIME).then(function (c) {
      return c.match(req).then(function (hit) {
        var net = fetch(req).then(function (res) {
          if (res.ok || res.type === 'opaque') c.put(req, res.clone());
          return res;
        }).catch(function () { return hit; });
        return hit || net;
      });
    }));
  }
});
