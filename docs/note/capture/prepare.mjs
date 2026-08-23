#!/usr/bin/env node
/**
 * 撮影用の写しを作る。
 *
 *   node docs/note/capture/prepare.mjs .capture [--keep-fonts]
 *
 * アプリの HTML と JavaScript には手を入れない。差し替えるのは
 * 「どこから読むか」だけ。
 *
 *   1. Google Fonts の読み込み先を、npm から取った同じ書体のローカル写しに向ける
 *      （作業環境から fonts.googleapis.com に出られないため。
 *        出られる環境なら --keep-fonts を渡して、この差し替えを飛ばす）
 *   2. CSP の connect-src に、手もとの PeerServer を足す
 *      （PeerJS の接続先は shots.mjs 側で差し替える）
 *
 * 書体を npm から入れておくこと:
 *   npm i --no-save @fontsource/shippori-mincho @fontsource/noto-sans-jp \
 *     @fontsource/yomogi @fontsource/kiwi-maru @fontsource/biz-udpgothic
 */
import { cp, mkdir, readFile, writeFile, rm, access } from 'node:fs/promises';
import path from 'node:path';

const ROOT = path.resolve(path.dirname(new URL(import.meta.url).pathname), '../../..');
const OUT = path.resolve(process.argv[2] || '.capture');
const KEEP_FONTS = process.argv.includes('--keep-fonts');

const FAMILIES = [
  ['Shippori Mincho', 'shippori-mincho', [400, 700]],
  ['Noto Sans JP', 'noto-sans-jp', [400, 700]],
  ['Yomogi', 'yomogi', [400]],
  ['Kiwi Maru', 'kiwi-maru', [400]],
  ['BIZ UDPGothic', 'biz-udpgothic', [400, 700]],
];

const FONT_LINK = '  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Shippori+Mincho:wght@400;700&family=Noto+Sans+JP:wght@400;700&family=Yomogi&family=Kiwi+Maru:wght@400;500&family=BIZ+UDPGothic:wght@400;700&display=swap">';

const exists = async (p) => access(p).then(() => true, () => false);

await rm(OUT, { recursive: true, force: true });
await mkdir(path.join(OUT, 'localfonts'), { recursive: true });

for (const name of ['index.html', 'offline.html', 'manifest.webmanifest', 'sw.js', 'favicon.png', 'icons', 'vendor']) {
  await cp(path.join(ROOT, name), path.join(OUT, name), { recursive: true });
}

let html = await readFile(path.join(OUT, 'index.html'), 'utf8');

if (!KEEP_FONTS) {
  const rules = [];
  for (const [family, slug, weights] of FAMILIES) {
    for (const weight of weights) {
      for (const subset of ['japanese', 'latin']) {
        const file = `${slug}-${subset}-${weight}-normal.woff2`;
        const src = path.join(ROOT, 'node_modules', '@fontsource', slug, 'files', file);
        if (!(await exists(src))) continue;
        await cp(src, path.join(OUT, 'localfonts', file));
        rules.push(`@font-face {\n  font-family: '${family}';\n  font-style: normal;\n  font-weight: ${weight};\n  font-display: swap;\n  src: url('localfonts/${file}') format('woff2');\n}`);
      }
    }
  }
  if (!rules.length) {
    console.error('❌ @fontsource の書体が見つかりません。npm i --no-save @fontsource/... を先に実行してください。');
    process.exit(1);
  }
  await writeFile(path.join(OUT, 'localfonts.css'), rules.join('\n') + '\n');
  if (!html.includes(FONT_LINK)) {
    console.error('❌ index.html の Google Fonts の <link> が見つかりません。書き方が変わったなら prepare.mjs も直してください。');
    process.exit(1);
  }
  html = html.replace(FONT_LINK, '  <link rel="stylesheet" href="localfonts.css">');
  html = html.replace("font-src 'self' https://fonts.gstatic.com;", "font-src 'self';");
  console.log(`書体を写した: ${rules.length} 面`);
}

const CONNECT = "connect-src 'self' data: blob: https://0.peerjs.com wss://0.peerjs.com;";
if (!html.includes(CONNECT)) {
  console.error('❌ index.html の connect-src が見つかりません。CSP の書き方が変わったなら prepare.mjs も直してください。');
  process.exit(1);
}
html = html.replace(CONNECT, "connect-src 'self' data: blob: ws://127.0.0.1:9000 http://127.0.0.1:9000 https://0.peerjs.com wss://0.peerjs.com;");

await writeFile(path.join(OUT, 'index.html'), html);
console.log(`撮影用の写しを作りました: ${OUT}`);
