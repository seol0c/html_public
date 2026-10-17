/* seolcount 웹판 오프라인 캐시.
   한 번 열어 둔 뒤에는 행사장에서 인터넷이 끊겨도 페이지가 뜬다.
   네트워크 먼저(그래야 고친 판이 바로 반영된다), 4초 안에 안 오면 캐시. 시트 전송(다른 도메인)은 건드리지 않는다. */
const C = 'seolcount-v1';
const FILES = ['./', './index.html', './manifest.webmanifest', './seolcount.gs',
               './icon-192.png', './icon-512.png', './apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(C).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== C).map(k => caches.delete(k))))
              .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (u.origin !== location.origin || e.request.method !== 'GET') return;
  e.respondWith((async () => {
    try {
      const r = await Promise.race([
        fetch(e.request),
        new Promise((_, no) => setTimeout(() => no(new Error('timeout')), 4000))
      ]);
      if (r && r.ok) { const cp = r.clone(); caches.open(C).then(c => c.put(e.request, cp)); }
      return r;
    } catch (err) {
      const hit = await caches.match(e.request, { ignoreSearch:true });
      return hit || fetch(e.request);
    }
  })());
});
