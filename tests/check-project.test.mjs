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
import { inspect, checkFiles, stripComments } from '../scripts/check-project.mjs';

const ids = (files) => inspect(files).map(p => p.id);
const hit = (files, id) => assert.ok(ids(files).includes(id), `${id} を検出できていない`);
const miss = (files, id) => assert.ok(!ids(files).includes(id), `${id} を誤検知している`);

test('秘密情報', () => {
  hit({ 'a.gs': 'const k = "AIzaSyA1234567890123456789012345678901234";' }, 'SECRET_APIKEY');
  hit({ 'a.gs': 'SpreadsheetApp.openById("1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789");' }, 'SECRET_ID');
  // PropertiesService 経由なら通す
  miss({ 'a.gs': 'SpreadsheetApp.openById(props.getProperty("SHEET_ID"));' }, 'SECRET_ID');
  hit({ 'a.gs': 'const to = "sensei@school.ed.jp";' }, 'SECRET_MAIL');
  // URL の中の文字列をメールと読み違えない
  miss({ 'a.html': '<a href="https://lh3.googleusercontent.com/d/abc">x</a>' }, 'SECRET_MAIL');
});

test('依存', () => {
  hit({ 'a.html': '<script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>' }, 'DEP_BABEL');
  hit({ 'a.html': '<script src="https://cdn.tailwindcss.com"></script>' }, 'DEP_TAILWIND_CDN');
  hit({ 'a.html': '<script src="https://cdnjs.cloudflare.com/ajax/libs/qrious/4.0.2/qrious.min.js"></script>' }, 'DEP_CDN_SCRIPT');
  // 自分側に持っていれば通す
  miss({ 'a.html': '<script>/* 同梱 */</script>' }, 'DEP_CDN_SCRIPT');
  hit({ 'a.html': '<link rel="stylesheet" href="https://unpkg.com/@picocss/pico/css/pico.min.css">' }, 'DEP_UNPINNED');
  miss({ 'a.html': '<link rel="stylesheet" href="https://unpkg.com/@picocss/pico@1.5.10/css/pico.min.css">' }, 'DEP_UNPINNED');
  // フォントは見た目だけなので版を固定しなくてよい
  miss({ 'a.html': '<link href="https://fonts.googleapis.com/css2?family=X&display=swap" rel="stylesheet">' }, 'DEP_UNPINNED');
});

test('viewport', () => {
  hit({ 'a.html': '<meta name="viewport" content="width=device-width, user-scalable=no">' }, 'VIEWPORT_NOZOOM');
  hit({ 'a.html': '<meta name="viewport" content="width=device-width, initial-scale=1.0">' }, 'VIEWPORT_NO_FIT');
  miss({ 'a.html': '<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">' }, 'VIEWPORT_NO_FIT');
  // GAS はサーバー側にもある。両方見る（v5 §5）
  hit({ 'a.gs': ".addMetaTag('viewport', 'width=device-width, initial-scale=1')" }, 'VIEWPORT_NO_FIT');
  miss({ 'a.gs': ".addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')" }, 'VIEWPORT_NO_FIT');
});

test('100vh', () => {
  hit({ 'a.html': '<style>.x { height: 100vh; }</style>' }, 'VIEWPORT_100VH');
  // @supports のフォールバックの中は正しい形なので通す（v5 §P4 の誤検知例）
  miss({ 'a.html': '<style>.x{height:100dvh}@supports not (height: 100dvh){.x{height:100vh}}</style>' }, 'VIEWPORT_100VH');
});

test('ふりがなの色', () => {
  // 色のついた面の上で決め打ちしている
  hit({ 'a.html': '<style>.tag-radio:checked + .tag-label rt { color: rgba(255,255,255,0.9); }</style>' }, 'RUBY_HARDCODED');
  hit({ 'a.html': '<style>button rt { color: #666; }</style>' }, 'RUBY_HARDCODED');
  // 継がせているなら正しい
  miss({ 'a.html': '<style>button rt, .tag-label rt { color: inherit; }</style>' }, 'RUBY_HARDCODED');
  // 白地の既定値は必要なので咎めない
  miss({ 'a.html': '<style>rt { color: #5f6368; }</style>' }, 'RUBY_HARDCODED');
  // <style> の外にある「CSS に見える文字列」を規則と読み違えない。
  // 実際に、JavaScript が組み立てる HTML の style 属性で誤検知した
  miss({ 'a.html': '<style>button rt{color:inherit}</style><script>e.innerHTML=\'<div style="color:#c0221f;">x</div>\'</script>' }, 'RUBY_HARDCODED');
});

test('Service Worker', () => {
  hit({ 'sw.js': 'caches.keys().then(ks => Promise.all(ks.map(k => caches.delete(k))))' }, 'SW_CACHE_WIPE');
  // 削除式を正規表現で追うと (k) => caches.delete(k) を見落とす。
  // 「startsWith で絞る式があるか」で見ているので、これは通る（v5 §P4 の実例）
  miss({ 'sw.js': 'const ks = await caches.keys(); ks.filter(k => k.startsWith(P)).map((k) => caches.delete(k))' }, 'SW_CACHE_WIPE');
  hit({ 'sw.js': 'localStorage.setItem("a", 1)' }, 'SW_LOCALSTORAGE');
  // 「localStorage は操作しない」という注意書きに反応しない（v5 §P4 の実例）
  miss({ 'sw.js': '/* Service Worker は localStorage を一切操作しない */\nself.addEventListener("fetch", () => {});' }, 'SW_LOCALSTORAGE');
});

test('禁止事項', () => {
  hit({ 'a.html': '<script>localStorage.clear()</script>' }, 'BAN_LS_CLEAR');
  hit({ 'a.html': '<script>w.postMessage(d, "*")</script>' }, 'BAN_POSTMESSAGE');
  miss({ 'a.html': '<script>w.postMessage(d, "https://example.com")</script>' }, 'BAN_POSTMESSAGE');
  hit({ 'appsscript.json': '{"oauthScopes":["https://www.googleapis.com/auth/drive"]}' }, 'BAN_SCOPE');
  miss({ 'appsscript.json': '{"oauthScopes":["https://www.googleapis.com/auth/drive.file"]}' }, 'BAN_SCOPE');
});

test('大きさ', () => {
  hit({ 'big.html': 'x\n'.repeat(5001) }, 'SIZE_LINES');
  miss({ 'ok.html': 'x\n'.repeat(100) }, 'SIZE_LINES');
});

test('構文', () => {
  hit({ 'a.gs': 'function f( {' }, 'SYNTAX');
  miss({ 'a.gs': 'function f() { return 1; }' }, 'SYNTAX');
  hit({ 'a.html': '<script>if (x { }</script>' }, 'SYNTAX');
  miss({ 'a.html': '<script>const s = `a${b}c`;</script>' }, 'SYNTAX');
  // 外部を読むだけの <script src> は中身が無いので対象にしない
  miss({ 'a.html': '<script src="https://x/y.js"></script>' }, 'SYNTAX');
  // GAS のスクリプトレットは、サーバー側で埋まるものなので畳んでから見る
  miss({ 'a.html': '<script>const u = "<?= getUrl() ?>";</script>' }, 'SYNTAX');
  // コメントの中に書いた <script> の文字を、本物のタグと読み違えない。
  // 同梱ライブラリの説明コメントに「この <script> の中に貼り替え」と
  // 書いたところ、そこから拾ってしまった
  miss({ 'a.html': '<!-- この <script> の中に貼る -->\n<script>const a = 1;</script>' }, 'SYNTAX');
});

test('置かれているべきファイル', () => {
  const missing = checkFiles(() => false).map(p => p.file);
  assert.ok(missing.includes('LICENSE') && missing.includes('.gitignore'), '不足を検出できていない');
  assert.deepStrictEqual(checkFiles(() => true), [], '揃っているのに指摘している');
});

test('コメントは判定前に落とす', () => {
  assert.strictEqual(stripComments('<!-- x -->a', 'html').trim(), 'a');
  assert.strictEqual(stripComments('/* x */a', 'js').trim(), 'a');
  // HTML では行コメント扱いをしない（URL の // を壊すため）
  assert.ok(stripComments('<a href="https://x/y">', 'html').includes('https://x/y'));
});
