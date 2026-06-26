const CACHE = 'gastro-v17';
const ASSETS = ['./', './index.html', './admin.html', './manifest.json'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  if (e.request.url.includes('script.google.com')) return;
  if (e.request.url.includes('imgur.com')) return;
  if (e.request.url.includes('fonts.google')) return;
  e.respondWith(
    fetch(e.request, { redirect: 'follow' }).then(res => {
      const clone = res.clone();
      if (res.status === 200 && res.type === 'basic')
        caches.open(CACHE).then(c => c.put(e.request, clone));
      return res;
    }).catch(() => caches.match(e.request))
  );
});
