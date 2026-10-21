const CACHE = 'voca-shell-v62';
const SHELL = ['./index_voca.html', './manifest_voca.json', './icon_voca.svg', './img/tm_hero.jpg', './img/univ/chonnam.png', './img/univ/chungbuk.png', './img/univ/cnu.png', './img/univ/dku.png', './img/univ/donga.png', './img/univ/ewha.png', './img/univ/gachon.png', './img/univ/gnu.png', './img/univ/hallym.png', './img/univ/jeju.png', './img/univ/jeonbuk.png', './img/univ/jeonju.png', './img/univ/kangwon.png', './img/univ/keimyung.png', './img/univ/konkuk.png', './img/univ/korea.png', './img/univ/kunsan.png', './img/univ/kyonggi.png', './img/univ/kyunghee.png', './img/univ/kyungpook.png', './img/univ/pusan.png', './img/univ/skku.png', './img/univ/snu.png', './img/univ/hanyang.png', './img/univ/sookmyung.png', './img/univ/wonkwang.png', './img/univ/woosuk.png', './img/univ/yeungnam.png', './img/univ/yonsei.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k.startsWith('voca-shell-') && k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// 앱 셸은 네트워크 우선, 실패 시 캐시. 다른 출처(GAS 등)와 POST 는 건드리지 않는다.
self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;
  e.respondWith(
    Promise.race([fetch(req), new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 4000))]).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
      return res;
    }).catch(() => caches.match(req).then(r => r || caches.match('./index_voca.html')))
  );
});
