/* デジタル・クラス新聞社 — Service Worker */
/*
 * 【最重要】activate では自アプリ以外のキャッシュを削除しない。
 *   同じオリジンに他の GIGA アプリが同居することがある。ここで
 *   caches.keys() の結果を全部消すと、このアプリを開くたびに
 *   同じ端末に入っている他のアプリのキャッシュまで巻き添えで消え、
 *   それらがオフラインで起動しなくなる。
 *   CACHE_PREFIX で始まるキャッシュだけを掃除する。
 *
 * この Service Worker は localStorage を一切さわらない。
 *   児童の記事（dnp_*）にも、他のアプリのキーにも触れない。
 */
const CACHE_PREFIX = 'digital-newspaper-';
// APP_VERSION は手で上げない。node tools/build-sw.mjs が先読み対象の中身から自動で決める
const APP_VERSION = 'vfcccd457'; /* __APP_VERSION__ */
const CACHE_NAME = CACHE_PREFIX + APP_VERSION;

const APP_SHELL = [
  // 書体そのもの（woff2）は先読みに入れない。入れると先読みが 1MB を超え、
  // 校内 Wi-Fi で 40 台が同時に開いたときに初回表示が止まる。
  // 画面が出れば必ず取りにいくので、その 1 回で下の実行時キャッシュに入る。
  './fonts.css',
  './',
  './index.html',
  './offline.html',
  './manifest.webmanifest',
  './favicon.png',
  './vendor/peerjs-1.5.2.min.js',
  './vendor/qr-creator-1.0.0.min.js',
  // 利用規約・プライバシーの行き先を出す部品。先読みに入れておかないと、
  // オフラインで開いたときだけフッターのリンクが 1 本も出ない
  // （行き先そのものは開けなくても、どこにあるかは見えているほうがいい）。
  './web/giga-app-links.js',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/maskable-192.png',
  './icons/maskable-512.png',
  './icons/apple-touch-icon.png'
];

// 実行時にキャッシュしてよい外部ホスト。フォントだけ。
// 実行コードは1バイトも外から取らないので、ここにスクリプトのホストは並ばない。
const RUNTIME_CACHE_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    // cache.addAll は1本でも失敗すると全体が巻き戻り、
    // オフライン起動そのものができなくなる。1本ずつ入れて取りこぼしを許容する。
    await Promise.all(APP_SHELL.map((url) =>
      cache.add(new Request(url, { cache: 'reload' }))
        .catch((err) => console.warn('[sw] precache skipped', url, err))
    ));
    // ここで skipWaiting() は呼ばない。
    // 新しい版はいったん待機させ、画面の帯で「さいしんに する」を
    // 押したときに初めて入れ替える。黙って入れ替えると、記事を書いている
    // 最中に中身が変わって驚かせてしまう。
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys
        // ← 自アプリ接頭辞のものだけを削除する。ここを外すと
        //    同一オリジンの他アプリを巻き添えにする。
        .filter((k) => k.startsWith(CACHE_PREFIX) && k !== CACHE_NAME)
        .map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);
  const isSameOrigin = url.origin === self.location.origin;
  const isRuntimeHost = RUNTIME_CACHE_HOSTS.includes(url.hostname);
  // PeerJS のシグナリングは動的な通信。キャッシュしない。
  if (!isSameOrigin && !isRuntimeHost) return;

  // ページ本体はネットワーク優先（更新を確実に取得）、オフライン時はキャッシュへ
  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request)
        .then((response) => {
          // エラーページ(404/5xx)をオフライン用に焼き込まないよう ok のときだけ保存する
          if (response && response.ok) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put('./index.html', copy));
          }
          return response;
        })
        // index.html すら入っていない（初回起動が圏外だった等）ときに
        // ブラウザ既定の白い画面を見せないよう offline.html を出す。
        .catch(async () =>
          (await caches.match('./index.html')) || (await caches.match('./offline.html'))
        )
    );
    return;
  }

  // 静的なものはキャッシュ優先＋裏で更新（stale-while-revalidate）
  event.respondWith(
    caches.match(request).then((cached) => {
      const networkFetch = fetch(request)
        .then((response) => {
          if (response && (response.ok || response.type === 'opaque')) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
          }
          return response;
        })
        .catch(() => cached);
      return cached || networkFetch;
    })
  );
});

// 画面の帯で「さいしんに する」が押されたときだけ待機を解除する
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});
