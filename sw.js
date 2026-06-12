/* Service Worker — офлайн-режим приложения «Абоненты» */
const CACHE = 'abonenty-v1';
const ASSETS = [
    './',
    'index.html',
    'css/style.css',
    'js/app.js',
    'js/vendor/exceljs.min.js',
    'manifest.webmanifest',
    'icons/icon-192.png',
    'icons/icon-512.png',
    'icons/icon-512-maskable.png',
];

self.addEventListener('install', (e) => {
    e.waitUntil(
        caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
    );
});

self.addEventListener('activate', (e) => {
    e.waitUntil(
        caches.keys()
            .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
            .then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', (e) => {
    const req = e.request;
    if (req.method !== 'GET') return;
    e.respondWith(
        caches.match(req).then(cached => {
            if (cached) return cached;
            return fetch(req).then(res => {
                const url = new URL(req.url);
                if (url.origin === location.origin && res.ok) {
                    const copy = res.clone();
                    caches.open(CACHE).then(c => c.put(req, copy));
                }
                return res;
            }).catch(() => {
                // Офлайн-фолбэк для навигации
                if (req.mode === 'navigate') return caches.match('index.html');
            });
        })
    );
});
