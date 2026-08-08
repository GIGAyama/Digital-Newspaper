#!/usr/bin/env node
/**
 * 品質ゲート（GIGA Standard v5）
 *
 * このリポジトリは GAS ウェブアプリ（C型）で、ビルドも npm 依存も無い。
 * そのため「静的に読めば必ず分かること」だけを検査する。
 * コントラストやタップ領域は実ブラウザでしか測れないので、ここでは扱わない
 * （測り方と実測値は AUDIT.md にある）。
 *
 *   node scripts/check-project.mjs           … 検査する
 *   node scripts/check-project.mjs --list    … 検査項目を出す
 *
 * 検査そのものが正しく動くかは tests/check-project.test.mjs で確かめている。
 * 「0件でした」だけでは、検査が動いているのか何も見ていないのか区別できないため、
 * わざと壊した入力を与えて、ちゃんと落ちることを見ている（v5 §P4）。
 */
import { readFileSync, existsSync, readdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Script } from 'node:vm';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/** コメントを落とす。注意書きに反応して誤検知しないようにするため（v5 §P4 の実例） */
export function stripComments(src, kind) {
  let s = src.replace(/<!--[\s\S]*?-->/g, '');          // HTML
  s = s.replace(/\/\*[\s\S]*?\*\//g, '');                // ブロック
  if (kind !== 'html') s = s.replace(/^\s*\/\/.*$/gm, ''); // 行（HTML は URL の // を壊すので除外）
  return s;
}

/**
 * 検査の本体。ファイルの中身を渡すと、見つかった問題を返す。
 * ファイル読み込みと分けてあるのは、テストからわざと壊した中身を渡せるようにするため。
 */
export function inspect(files) {
  const found = [];
  const add = (id, file, message) => found.push({ id, file, message });
  const html = Object.entries(files).filter(([f]) => f.endsWith('.html'));
  const gs = Object.entries(files).filter(([f]) => f.endsWith('.gs'));

  // --- 秘密情報（v5 A2 / Part IV） -------------------------------------
  for (const [f, raw] of Object.entries(files)) {
    const src = stripComments(raw, f.endsWith('.html') ? 'html' : 'js');
    // Google の API キーは AIza で始まり 39 文字。末尾に \b を置くと
    // 長めのダミー文字列を取りこぼすので、始まりだけを見る
    if (/\bAIza[0-9A-Za-z_-]{35}/.test(src)) add('SECRET_APIKEY', f, 'API キーらしき文字列が直書きされている');
    // スプレッドシートID / フォルダID の直書き（PropertiesService を使うこと）
    const idLit = src.match(/openById\(\s*['"][^'"]{20,}['"]|getFolderById\(\s*['"][^'"]{20,}['"]/);
    if (idLit) add('SECRET_ID', f, 'シートID／フォルダIDが直書きされている: ' + idLit[0].slice(0, 40));
    if (/[\w.+-]+@(?!example\.)[\w-]+\.[\w.]+/.test(src.replace(/https?:\/\/\S+/g, '')))
      add('SECRET_MAIL', f, 'メールアドレスらしき文字列が直書きされている');
  }

  // --- 依存（v5 §1 / §6 / B6） ----------------------------------------
  for (const [f, raw] of html) {
    const src = stripComments(raw, 'html');
    if (/babel\/standalone/.test(src)) add('DEP_BABEL', f, 'ブラウザへ @babel/standalone を送っている');
    if (/cdn\.tailwindcss\.com/.test(src)) add('DEP_TAILWIND_CDN', f, 'ブラウザ内で CSS を生成する Tailwind CDN を使っている');
    // 外部から取る「実行コード」は 0 バイトが目標。塞がれると機能が黙って壊れる
    for (const m of src.matchAll(/<script[^>]+src\s*=\s*["'](https?:)?\/\/([^"']+)["'][^>]*>/gi))
      add('DEP_CDN_SCRIPT', f, '実行コードを外部から読んでいる: ' + m[2].split('/')[0]);
    // 残る外部資産（スタイル）は版の固定を要求する
    for (const m of src.matchAll(/<link[^>]+href\s*=\s*["']https?:\/\/([^"']+)["'][^>]*>/gi)) {
      const url = m[1];
      if (/fonts\.googleapis\.com|fonts\.gstatic\.com/.test(url)) continue;  // フォントは見た目だけ
      if (!/@\d+\.\d+\.\d+|\/\d+\.\d+\.\d+\//.test(url))
        add('DEP_UNPINNED', f, '外部スタイルの版が固定されていない: ' + url.split('/')[0]);
    }
  }

  // --- 表示（v5 §2） ---------------------------------------------------
  for (const [f, raw] of [...html, ...gs]) {
    const src = stripComments(raw, f.endsWith('.html') ? 'html' : 'js');
    if (/user-scalable\s*=\s*no|maximum-scale\s*=\s*1/.test(src))
      add('VIEWPORT_NOZOOM', f, '拡大を禁止している（見えづらい子が拡大できなくなる）');

    // viewport は HTML の <meta> と GAS の addMetaTag の両方にあり得る。両方に要る（v5 §5）
    for (const m of src.matchAll(/<meta[^>]+name\s*=\s*["']viewport["'][^>]+content\s*=\s*["']([^"']+)["']/gi))
      if (!/viewport-fit\s*=\s*cover/.test(m[1])) add('VIEWPORT_NO_FIT', f, '<meta> の viewport に viewport-fit=cover が無い');
    for (const m of src.matchAll(/addMetaTag\(\s*['"]viewport['"]\s*,\s*['"]([^'"]+)['"]/g))
      if (!/viewport-fit\s*=\s*cover/.test(m[1])) add('VIEWPORT_NO_FIT', f, 'addMetaTag の viewport に viewport-fit=cover が無い');

    // 100vh 単独は使わない。@supports のフォールバックの中にあるものは正しい（v5 §P4 の誤検知例）
    const noSupports = src.replace(/@supports\s+not\s*\(height:\s*100dvh\)\s*\{[\s\S]*?\}\s*\}?/g, '');
    if (/height:\s*100vh/.test(noSupports)) add('VIEWPORT_100VH', f, '100vh を単独で使っている（dvh を使うこと）');
  }

  // --- ふりがな（v5 §4） ------------------------------------------------
  for (const [f, raw] of html) {
    // CSS の規則として読むので、<style> の中だけを対象にする。
    // ファイル全体に当てると、JavaScript の中の HTML 文字列
    // （'<div style="color:#c0221f;">…' など）を規則と読み違える。
    const src = [...stripComments(raw, 'html').matchAll(/<style[^>]*>([\s\S]*?)<\/style>/gi)]
      .map(m => m[1]).join('\n');
    // 色のついた面の上に重なる rt に、色を決め打ちしていないか。
    // 「面の上では継がせる」形（color: inherit）は正しいので通す。
    for (const m of src.matchAll(/([^{}]*\brt\b[^{}]*)\{([^}]*)\}/g)) {
      const sel = m[1].trim(), body = m[2];
      const color = body.match(/(?:^|[;\s])color\s*:\s*([^;]+)/);
      if (!color) continue;
      const v = color[1].trim();
      if (v === 'inherit' || v === 'currentColor') continue;
      // 「色のついた面」を指すセレクタの中で決め打ちしているものだけを咎める。
      // 白地の既定値（rt { color: ... }）は必要なので通す。
      if (/:checked|\[class\*?=|\bbutton\b|\bbtn\b|\.badge|\.tag-label|bg-/.test(sel))
        add('RUBY_HARDCODED', f, '色のついた面の rt に色を決め打ちしている: ' + sel.slice(0, 60) + ' { color: ' + v + ' }');
    }
  }

  // --- Service Worker（このリポジトリには無いが、C+型へ進んだときのため） ---
  for (const [f, raw] of Object.entries(files)) {
    if (!/sw\.js$/.test(f)) continue;
    const src = stripComments(raw, 'js');
    // 「消す式」ではなく「startsWith で絞る式があるか」を見る（v5 §P4 の実例）
    if (/caches\.keys\(\)/.test(src) && !/startsWith/.test(src))
      add('SW_CACHE_WIPE', f, '自アプリ以外のキャッシュまで消している疑い（startsWith で絞ること）');
    if (/localStorage/.test(src)) add('SW_LOCALSTORAGE', f, 'Service Worker が localStorage に触れている');
  }

  // --- 禁止事項（Part IV） ---------------------------------------------
  for (const [f, raw] of Object.entries(files)) {
    const src = stripComments(raw, f.endsWith('.html') ? 'html' : 'js');
    if (/localStorage\.clear\(\)/.test(src)) add('BAN_LS_CLEAR', f, 'localStorage.clear() を使っている');
    if (/postMessage\([^)]*,\s*['"]\*['"]\s*\)/.test(src)) add('BAN_POSTMESSAGE', f, 'postMessage の宛先が * になっている');
    if (/auth\/drive['"]|https:\/\/mail\.google\.com\//.test(src)) add('BAN_SCOPE', f, '広すぎる OAuth スコープを要求している');
  }

  // --- 構文 -------------------------------------------------------------
  // GAS のコードは実行して確かめられない。せめて構文だけは見る。
  // .gs は V8 ランタイムの JavaScript なので、そのまま構文解析できる。
  // HTML の中のインラインの <script> も見る（GAS はコードを全部そこに持つため）。
  const parse = (code, f, where) => {
    try { new Script(code); }
    catch (e) { add('SYNTAX', f, `${where}に構文エラー: ${e.message}`); }
  };
  for (const [f, raw] of gs) parse(raw, f, 'ファイル');
  for (const [f, raw] of html) {
    // コメントを先に落とす。落とさないと、コメントの中に書いた <script> の
    // 文字を本物のタグと読み違え、そこからコメントの残りごと拾ってしまう。
    // （実際に、同梱ライブラリの説明コメントで踏んだ）
    for (const m of stripComments(raw, 'html').matchAll(/<script(?![^>]*\bsrc\s*=)[^>]*>([\s\S]*?)<\/script>/gi)) {
      // スクリプトレットが混ざっていると解析できない。GAS がサーバー側で
      // 埋めるものなので、ここでは畳んでから見る。
      // 0 に置き換えるのは、文字列の中（"<?= x ?>" → "0"）でも
      // 外（var n = <?= x ?>; → var n = 0;）でも構文が壊れないため。
      parse(m[1].replace(/<\?=?!?[\s\S]*?\?>/g, '0'), f, 'インラインの <script> ');
    }
  }

  // --- 大きさ（v5 §8） --------------------------------------------------
  for (const [f, raw] of Object.entries(files)) {
    const lines = raw.split('\n').length;
    if (lines > 5000) add('SIZE_LINES', f, `${lines} 行（5,000 行を超えている）`);
    if (Buffer.byteLength(raw) > 400 * 1024) add('SIZE_BYTES', f, `${Math.round(Buffer.byteLength(raw) / 1024)}KB（400KB を超えている）`);
  }

  return found;
}

/** リポジトリに置かれているべきファイル */
export function checkFiles(exists) {
  const missing = [];
  for (const [f, why] of [
    ['LICENSE', '配布条件が示されない'],
    ['.gitignore', '.clasp.json や .env が混入する'],
    ['README.md', '開発者向けの説明が無い'],
    ['MANUAL.md', '先生向けの説明が無い'],
    ['AUDIT.md', '実測の記録が無い'],
    ['.github/dependabot.yml', '依存の更新が来ない'],
  ]) if (!exists(f)) missing.push({ id: 'FILE_MISSING', file: f, message: why });
  return missing;
}

// --- ここから下は実行用（テストからは import されない） -------------------
if (import.meta.url === `file://${process.argv[1]}`) {
  if (process.argv.includes('--list')) {
    console.log([
      'SECRET_APIKEY / SECRET_ID / SECRET_MAIL … 秘密情報の直書き',
      'DEP_BABEL / DEP_TAILWIND_CDN / DEP_CDN_SCRIPT / DEP_UNPINNED … 依存',
      'VIEWPORT_NOZOOM / VIEWPORT_NO_FIT / VIEWPORT_100VH … 表示',
      'RUBY_HARDCODED … ふりがなの色',
      'SW_CACHE_WIPE / SW_LOCALSTORAGE … Service Worker',
      'BAN_LS_CLEAR / BAN_POSTMESSAGE / BAN_SCOPE … 禁止事項',
      'SIZE_LINES / SIZE_BYTES … 大きさ',
      'FILE_MISSING … 置かれているべきファイル',
    ].join('\n'));
    process.exit(0);
  }

  const files = {};
  const walk = (dir, rel = '') => {
    for (const e of readdirSync(join(ROOT, dir), { withFileTypes: true })) {
      if (e.name === '.git' || e.name === 'node_modules') continue;
      const r = rel ? `${rel}/${e.name}` : e.name;
      if (e.isDirectory()) walk(join(dir, e.name), r);
      else if (/\.(html|gs|js|mjs|json)$/.test(e.name) && !r.startsWith('scripts/') && !r.startsWith('tests/'))
        files[r] = readFileSync(join(ROOT, dir, e.name), 'utf8');
    }
  };
  walk('.');

  const found = [...checkFiles(f => existsSync(join(ROOT, f))), ...inspect(files)];

  console.log(`検査したファイル: ${Object.keys(files).length}本`);
  if (found.length === 0) {
    console.log('✅ 指摘なし');
    process.exit(0);
  }
  for (const p of found) console.log(`❌ [${p.id}] ${p.file}: ${p.message}`);
  console.log(`\n${found.length}件`);
  process.exit(1);
}
