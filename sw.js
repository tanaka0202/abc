// オフラインでも画面を開けるようにする（アプリ本体のファイルだけを保存。会議データは保存しない）
const CACHE = 'mtgrec-v2.2.0';
const FILES = [
  './', './index.html', './auth.html', './manifest.webmanifest', './css/app.css', './icons/icon.svg', './icons/icon-192.png', './icons/icon-512.png',
  './js/app.js', './js/util.js', './js/db.js', './js/settings.js', './js/wav.js', './js/textproc.js', './js/audio-worklet.js',
  './js/recorder.js', './js/auth.js', './js/auth-callback.js', './js/onedrive.js', './js/gcal.js', './js/gemini.js',
  './js/usage.js', './js/vocab.js', './js/sync.js', './js/memos.js', './js/pipeline.js', './js/ui.js', './js/qr.js', './js/setuplink.js',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES.map((f) => new Request(f, { cache: 'no-cache' })))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

// 同じサイトのファイルだけ：まずネットから取り、失敗したら保存済みのものを使う（更新がすぐ反映される）
self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  if (url.pathname.endsWith('/auth.html')) return; // サインインの戻り先は常にネットから
  e.respondWith(
    // cache: 'no-cache' … ブラウザ側に残った古いファイル（GitHub Pagesは10分間の保持を指定）を使わず、毎回サーバーに更新を確かめる
    fetch(e.request, { cache: 'no-cache' }).then((res) => {
      if (res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(e.request, copy)); }
      return res;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match('./index.html'))),
  );
});
