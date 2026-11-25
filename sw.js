// Network first, cache fallback: the app still opens in the storage room with no signal.
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || e.request.url.includes('script.google')) return;
  e.respondWith(fetch(e.request).then(r => {
    const c = r.clone(); caches.open('v1').then(k => k.put(e.request, c)); return r;
  }).catch(() => caches.match(e.request)));
});
