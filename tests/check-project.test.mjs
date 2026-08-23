#!/usr/bin/env node
/**
 * 品質ゲートそのものを、わざと壊して確かめる（GIGA Standard v5 §P4）
 *
 * 「0件でした」だけでは、検査が動いているのか何も見ていないのか区別できない。
 * 各項目について「引っかかるべき入力」と「引っかかってはいけない入力」の
 * 両方を与え、期待どおりに動くことを見る。
 *
 *   node --test tests/
 */
import test from 'node:test';
import assert from 'node:assert';
import { inspect, checkFiles, checkAssets, stripComments } from '../scripts/check-project.mjs';

const ids = (files, opts) => inspect(files, opts).map((p) => p.id);
const hit = (files, id, opts) => assert.ok(ids(files, opts).includes(id), `${id} を検出できていない`);
const miss = (files, id, opts) => assert.ok(!ids(files, opts).includes(id), `${id} を誤検知している`);

/* アプリ本体だけに課す項目を確かめるための、最低限そろっている entry。
   ここに足りないものを1つずつ抜いて、抜いたぶんだけが赤くなることを見る。 */
const GOOD_ENTRY = `<!DOCTYPE html><html><head>
<script>addEventListener('beforeinstallprompt', function(e){ window.__deferredInstallPrompt = e; dispatchEvent(new Event('pwa-installable')); });</script>
<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'">
<link rel="apple-touch-icon" href="icons/apple-touch-icon.png">
<style>
  body { padding-bottom: env(safe-area-inset-bottom); touch-action: manipulation; font-size: clamp(15px, 1vw, 17px); }
  .btn { min-height: 44px; }
  :focus-visible { outline: 3px solid #1f4e8c; }
  @media (prefers-reduced-motion: reduce) { * { transition-duration: .01ms !important; } }
  @media print { .no-print { display: none; } }
</style>
</head><body>
<div id="updateBar">あたらしい バージョンが あります</div>
<script>
  if (window.visualViewport) { window.visualViewport.addEventListener('resize', function(){}); }
  addEventListener('pagehide', function(){});
  reg.waiting.postMessage({ type: 'SKIP_WAITING' });
</script>
</body></html>`;
const withEntry = (html) => ({ 'index.html': html });
const drop = (needle) => GOOD_ENTRY.replace(needle, '');

test('そろっている entry は、本体むけの項目で1つも赤くならない', () => {
  const bad = inspect(withEntry(GOOD_ENTRY)).map((p) => p.id);
  assert.deepStrictEqual(bad, [], '想定外の指摘: ' + bad.join(', '));
});

test('コメントの落としかた', () => {
  // accept="image/*" の /* をコメントの始まりと読み違えない。
  // 実際にこれで、ファイル後半の 17,000 字ぶんが消えて
  // 「更新のお知らせが無い」と誤判定していた。
  const html = '<input accept="image/*"><p>あたらしい バージョン</p>';
  assert.ok(stripComments(html, 'html').includes('あたらしい バージョン'));
  // <style> の中の CSS コメントは落とす
  assert.ok(!stripComments('<style>/* めも */ .a{}</style>', 'html').includes('めも'));
  // HTML のコメントは落とす
  assert.ok(!stripComments('<!-- めも --><p>本文</p>', 'html').includes('めも'));
  // URL の // を行コメントとして落とさない
  assert.ok(stripComments('<a href="https://example.com/x">y</a>', 'html').includes('example.com'));
  // JavaScript では行コメントも落とす
  assert.ok(!stripComments('// めも\nvar a = 1;', 'js').includes('めも'));
});

test('秘密情報', () => {
  hit({ 'a.js': 'const k = "AIzaSyA1234567890123456789012345678901234";' }, 'SECRET_APIKEY');
  miss({ 'a.js': 'const k = props.get("API_KEY");' }, 'SECRET_APIKEY');
  hit({ 'a.js': 'const to = "sensei@school.ed.jp";' }, 'SECRET_MAIL');
  // URL の中の文字列をメールと読み違えない
  miss({ 'a.html': '<a href="https://lh3.googleusercontent.com/d/abc">x</a>' }, 'SECRET_MAIL');
  // npm の版指定をメールと読み違えない
  miss({ 'package.json': '{"scripts":{"x":"npm install @google/clasp@3.3.0"}}' }, 'SECRET_MAIL');
  // 数字を含むドメインは今までどおり拾う
  hit({ 'a.js': 'const to = "sensei@school2.ed.jp";' }, 'SECRET_MAIL');
  // 法務ページの問い合わせ先は「載っているべきもの」。秘密の漏えいではない
  miss({ 'privacy.html': '<p>お問い合わせ: madoguchi@school.ed.jp</p>' }, 'SECRET_MAIL');
  miss({ 'terms.html': '<p>お問い合わせ: madoguchi@school.ed.jp</p>' }, 'SECRET_MAIL');
  // 名前が似ているだけの別ファイルは、対象から外さない
  hit({ 'my-privacy.html': '<p>madoguchi@school.ed.jp</p>' }, 'SECRET_MAIL');
});

test('依存', () => {
  hit({ 'a.html': '<script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>' }, 'DEP_BABEL');
  hit({ 'a.html': '<script src="https://cdn.tailwindcss.com"></script>' }, 'DEP_TAILWIND_CDN');
  hit({ 'a.html': '<script src="https://unpkg.com/peerjs@1.5.2/dist/peerjs.min.js"></script>' }, 'DEP_CDN_SCRIPT');
  // 自分側に置いていれば通す（このリポジトリは vendor/ に同梱している）
  miss({ 'a.html': '<script src="vendor/peerjs-1.5.2.min.js"></script>' }, 'DEP_CDN_SCRIPT');
  // 版の固定されていない外部スタイルは拾う
  hit({ 'a.html': '<link rel="stylesheet" href="https://unpkg.com/some/pkg.css">' }, 'DEP_UNPINNED');
  miss({ 'a.html': '<link rel="stylesheet" href="https://unpkg.com/@picocss/pico@1.5.10/css/pico.min.css">' }, 'DEP_UNPINNED');
  // rel と href の順が逆でも見る
  hit({ 'a.html': '<link href="https://unpkg.com/some/pkg.css" rel="stylesheet">' }, 'DEP_UNPINNED');
  // canonical / icon / manifest は「版を固定する」対象ではない
  miss({ 'a.html': '<link rel="canonical" href="https://digital-newspaper.giga-school.com/privacy.html">' }, 'DEP_UNPINNED');
  miss({ 'a.html': '<link rel="icon" href="https://example.com/favicon.png">' }, 'DEP_UNPINNED');
  // フォントは見た目だけなので通す
  miss({ 'a.html': '<link href="https://fonts.googleapis.com/css2?family=X&display=swap" rel="stylesheet">' }, 'DEP_UNPINNED');
});

test('viewport と 100vh', () => {
  hit({ 'a.html': '<meta name="viewport" content="width=device-width, user-scalable=no">' }, 'VIEWPORT_NOZOOM');
  hit({ 'a.html': '<meta name="viewport" content="width=device-width, initial-scale=1.0">' }, 'VIEWPORT_NO_FIT');
  miss({ 'a.html': '<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">' }, 'VIEWPORT_NO_FIT');
  hit({ 'a.html': '<style>.x { height: 100vh; }</style>' }, 'VIEWPORT_100VH');
  // dvh のフォールバックとして書いてあるものは正しい
  miss({ 'a.html': '<style>.x{height:100dvh}\n@supports not (height:100dvh){.x{height:100vh}}</style>' }, 'VIEWPORT_100VH');
});

test('表示への配慮（すべての HTML）', () => {
  hit({ 'a.html': '<style>body{margin:0}</style>' }, 'A11Y_NO_SAFE_AREA');
  miss({ 'a.html': '<style>body{padding:env(safe-area-inset-bottom)}</style>' }, 'A11Y_NO_SAFE_AREA');
  hit({ 'a.html': '<style>body{margin:0}</style>' }, 'A11Y_NO_REDUCED_MOTION');
  miss({ 'a.html': '<style>@media (prefers-reduced-motion: reduce){*{}}</style>' }, 'A11Y_NO_REDUCED_MOTION');
  hit({ 'a.html': '<style>body{margin:0}</style>' }, 'A11Y_NO_TOUCH_ACTION');
  miss({ 'a.html': '<style>body{touch-action:manipulation}</style>' }, 'A11Y_NO_TOUCH_ACTION');
  hit({ 'a.html': '<style>body{font-size:16px}</style>' }, 'A11Y_NO_FLUID_TYPE');
  miss({ 'a.html': '<style>body{font-size:clamp(15px,1vw,18px)}</style>' }, 'A11Y_NO_FLUID_TYPE');
});

test('アプリ本体にだけ課す項目', () => {
  hit(withEntry(drop('.btn { min-height: 44px; }')), 'A11Y_NO_TAP44');
  hit(withEntry(drop(':focus-visible { outline: 3px solid #1f4e8c; }')), 'A11Y_NO_FOCUS_VISIBLE');
  hit(withEntry(drop('@media print { .no-print { display: none; } }')), 'NO_PRINT_CSS');
  hit(withEntry(drop("if (window.visualViewport) { window.visualViewport.addEventListener('resize', function(){}); }")), 'NO_VISUAL_VIEWPORT');
  hit(withEntry(drop(`<meta http-equiv="Content-Security-Policy" content="default-src 'none'">`)), 'CSP_MISSING');
  hit(withEntry(drop("addEventListener('pagehide', function(){});")), 'NO_PAGEHIDE_FLUSH');
  // 別名の HTML には課さない（本体だけの決まり）
  miss({ 'other.html': '<p>x</p>' }, 'CSP_MISSING');
});

test('Canvas の画素密度', () => {
  hit(withEntry(GOOD_ENTRY + "<script>c.getContext('2d');</script>"), 'CANVAS_NO_DPR');
  miss(withEntry(GOOD_ENTRY + "<script>c.width = 50 * devicePixelRatio; c.getContext('2d');</script>"), 'CANVAS_NO_DPR');
  // Canvas を使っていなければ問わない
  miss(withEntry(GOOD_ENTRY), 'CANVAS_NO_DPR');
});

test('PWA', () => {
  hit(withEntry(drop('<div id="updateBar">あたらしい バージョンが あります</div>')), 'PWA_NO_UPDATE_NOTICE');
  hit(withEntry(drop("reg.waiting.postMessage({ type: 'SKIP_WAITING' });")), 'PWA_NO_SKIP_WAITING_UI');
  hit(withEntry(drop('<link rel="apple-touch-icon" href="icons/apple-touch-icon.png">')), 'PWA_NO_APPLE_ICON');
  hit(withEntry(GOOD_ENTRY.replace('__deferredInstallPrompt', 'somethingElse')), 'PWA_NO_INSTALL_BUTTON');
  // 合図の捕捉が外部スクリプトより後ろだと赤くする
  hit(withEntry('<head><script src="vendor/x.js"></script><script>addEventListener("beforeinstallprompt", function(){});</script></head>'), 'PWA_INSTALL_LATE');
  miss(withEntry(GOOD_ENTRY), 'PWA_INSTALL_LATE');
});

test('manifest', () => {
  const m = (o) => ({ 'manifest.webmanifest': JSON.stringify(o) });
  const icons = [
    { sizes: '192x192', purpose: 'any' }, { sizes: '512x512', purpose: 'any' },
    { sizes: '192x192', purpose: 'maskable' }, { sizes: '512x512', purpose: 'maskable' },
  ];
  // 独自ドメインの直下なら "./"。リポジトリ名の絶対パスに戻すと
  // scope がページの URL を含まなくなり、PWA として入れられなくなる
  miss(m({ id: './', scope: './', start_url: './?source=pwa', icons }), 'PWA_MANIFEST_PATH');
  hit(m({ id: '/Digital-Newspaper/', scope: '/Digital-Newspaper/', start_url: '/Digital-Newspaper/', icons }), 'PWA_MANIFEST_PATH');
  // オリジンを他アプリと共有する配置なら、その逆
  const opts = { hasCname: false, repoName: 'Digital-Newspaper' };
  miss(m({ id: '/Digital-Newspaper/', scope: '/Digital-Newspaper/', start_url: '/Digital-Newspaper/', icons }), 'PWA_MANIFEST_PATH', opts);
  hit(m({ id: './', scope: './', start_url: './', icons }), 'PWA_MANIFEST_PATH', opts);
  // アイコンの欠け
  hit(m({ id: './', scope: './', start_url: './', icons: icons.slice(0, 3) }), 'PWA_ICONS');
  hit({ 'manifest.webmanifest': '{ broken' }, 'PWA_MANIFEST_BROKEN');
});

test('Service Worker', () => {
  const base = "const APP_VERSION='v1'; caches.keys(); self.addEventListener('message', e => { if (e.data.type === 'SKIP_WAITING') self.skipWaiting(); }); const S=['./offline.html'];";
  miss({ 'sw.js': base.replace('caches.keys();', "caches.keys().then(k => k.filter(x => x.startsWith(P)));") }, 'SW_CACHE_WIPE');
  hit({ 'sw.js': base }, 'SW_CACHE_WIPE');
  hit({ 'sw.js': base + ' localStorage.getItem("x");' }, 'SW_LOCALSTORAGE');
  // 「localStorage には触れない」という注意書きだけで赤くしない
  miss({ 'sw.js': base + '\n/* localStorage には一切触れない */' }, 'SW_LOCALSTORAGE');
  // 写真の置き場所（IndexedDB）にも触らせない
  hit({ 'sw.js': base + ' indexedDB.deleteDatabase("dnp_photos_v1");' }, 'SW_INDEXEDDB');
  miss({ 'sw.js': base + '\n/* indexedDB には一切触れない */' }, 'SW_INDEXEDDB');
  miss({ 'sw.js': base }, 'SW_INDEXEDDB');
  hit({ 'sw.js': "const APP_VERSION='v1'; const S=['./offline.html'];" }, 'SW_NO_SKIP_WAITING');
  hit({ 'sw.js': "const APP_VERSION='v1'; SKIP_WAITING;" }, 'SW_NO_OFFLINE_PAGE');
  hit({ 'sw.js': "SKIP_WAITING; const S=['./offline.html'];" }, 'SW_NO_APP_VERSION');
});

test('ふりがなの色', () => {
  // 色のついた面の上で決め打ちしている
  hit({ 'a.html': '<style>.tag-radio:checked + .tag-label rt { color: rgba(255,255,255,0.9); }</style>' }, 'RUBY_HARDCODED');
  hit({ 'a.html': '<style>button rt { color: #666; }</style>' }, 'RUBY_HARDCODED');
  // 継がせる形は正しい
  miss({ 'a.html': '<style>button rt, .tag-label rt { color: inherit; }</style>' }, 'RUBY_HARDCODED');
  // 白地の既定値は必要なので通す
  miss({ 'a.html': '<style>rt { color: #5f6368; }</style>' }, 'RUBY_HARDCODED');
  // JavaScript の中の HTML 文字列を CSS の規則と読み違えない
  miss({ 'a.html': '<script>var s = ".btn rt { color: red }";</script></style>' }, 'RUBY_HARDCODED');
});

test('禁止事項', () => {
  hit({ 'a.js': 'localStorage.clear()' }, 'BAN_LS_CLEAR');
  miss({ 'a.js': 'Object.keys(localStorage).forEach(k => { if (k.startsWith(NS)) localStorage.removeItem(k); })' }, 'BAN_LS_CLEAR');
  hit({ 'a.js': 'w.postMessage(data, "*")' }, 'BAN_POSTMESSAGE');
  miss({ 'a.js': 'sw.postMessage({ type: "SKIP_WAITING" })' }, 'BAN_POSTMESSAGE');
});

test('構文', () => {
  hit({ 'a.html': '<script>function ( {</script>' }, 'SYNTAX');
  miss({ 'a.html': '<script>function f() { return 1; }</script>' }, 'SYNTAX');
  // 外部スクリプトの読み込みは中身が無いので対象外
  miss({ 'a.html': '<script src="vendor/x.js"></script>' }, 'SYNTAX');
});

test('大きさ', () => {
  hit({ 'a.html': 'x\n'.repeat(5001) }, 'SIZE_LINES');
  miss({ 'a.html': 'x\n'.repeat(100) }, 'SIZE_LINES');
  hit({ 'a.html': 'x'.repeat(400 * 1024 + 1) }, 'SIZE_BYTES');
});

test('画像の大きさ', () => {
  const only = (assets) => checkAssets(assets).map((p) => p.file);
  assert.deepStrictEqual(only([{ path: 'favicon.png', bytes: 40 * 1024 }]), ['favicon.png']);
  assert.deepStrictEqual(only([{ path: 'favicon.png', bytes: 20 * 1024 }]), []);
  assert.deepStrictEqual(only([{ path: 'icons/icon-512.png', bytes: 70 * 1024 }]), ['icons/icon-512.png']);
  // 512 のアイコンは 60KB。ふつうの画像の 150KB を先に当ててはいけない
  assert.deepStrictEqual(only([{ path: 'icons/icon-512.png', bytes: 100 * 1024 }]), ['icons/icon-512.png']);
  assert.deepStrictEqual(only([{ path: 'docs/x.png', bytes: 200 * 1024 }]), ['docs/x.png']);
  assert.deepStrictEqual(only([{ path: 'docs/x.png', bytes: 100 * 1024 }]), []);
});

test('置かれているべきファイル', () => {
  const all = () => true;
  assert.deepStrictEqual(checkFiles(all), []);
  const without = (name) => (f) => f !== name;
  assert.deepStrictEqual(checkFiles(without('sw.js')).map((p) => p.file), ['sw.js']);
  assert.deepStrictEqual(checkFiles(without('offline.html')).map((p) => p.file), ['offline.html']);
  assert.deepStrictEqual(checkFiles(without('AUDIT.md')).map((p) => p.file), ['AUDIT.md']);
});
