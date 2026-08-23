#!/usr/bin/env node
/**
 * 品質ゲート（GIGA Standard v5 / A型：ビルドの無い静的 PWA）
 *
 * このリポジトリは GitHub Pages に置くだけの静的アプリで、ビルドも npm 依存も無い。
 * そのため「静的に読めば必ず分かること」だけを検査する。
 * コントラストやタップ領域の実寸は実ブラウザでしか測れないので、ここでは扱わない
 * （測り方と実測値は AUDIT.md にある）。
 *
 *   node scripts/check-project.mjs           … 検査する
 *   node scripts/check-project.mjs --list    … 検査項目を出す
 *
 * 検査そのものが正しく動くかは tests/check-project.test.mjs で確かめている。
 * 「0件でした」だけでは、検査が動いているのか何も見ていないのか区別できないため、
 * わざと壊した入力を与えて、ちゃんと落ちることを見ている（v5 §P4）。
 *
 * 検査を緩めたいときはコードを消さず、quality.config.json の waived に
 * 「どの項目を、なぜ」を書いて明示的に許可すること。
 */
import { readFileSync, existsSync, readdirSync, statSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Script } from 'node:vm';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

const stripBlockComments = (s) => s.replace(/\/\*[\s\S]*?\*\//g, '');

/**
 * コメントを落とす。注意書きに反応して誤検知しないようにするため（v5 §P4 の実例）
 *
 * HTML では、CSS と JavaScript のコメントを <style> と <script> の中でしか
 * 落とさない。ファイル全体に当てると accept="image/*" の /* を
 * コメントの始まりと読み違え、そこから次に現れる閉じ記号までを丸ごと食う。
 * 実際にそれで、ファイルの後半にある更新のお知らせが「無い」と判定された
 * （17,000 字ぶんが消えていた）。
 */
export function stripComments(src, kind) {
  if (kind === 'html') {
    const noHtmlComments = src.replace(/<!--[\s\S]*?-->/g, '');
    return noHtmlComments.replace(
      /(<(style|script)\b[^>]*>)([\s\S]*?)(<\/\2>)/gi,
      (_all, open, _tag, body, close) => open + stripBlockComments(body) + close
    );
  }
  // 行コメントを落とすのは JavaScript のときだけ。HTML に当てると URL の // を壊す。
  return stripBlockComments(src).replace(/^\s*\/\/.*$/gm, '');
}

/**
 * 検査の本体。ファイルの中身を渡すと、見つかった問題を返す。
 * ファイル読み込みと分けてあるのは、テストからわざと壊した中身を渡せるようにするため。
 *
 * opts.entry    … アプリ本体の HTML（ここにだけ課す項目がある）
 * opts.hasCname … 独自ドメインの直下に置くか（manifest の id の期待値が変わる）
 */
export function inspect(files, opts = {}) {
  const entryName = opts.entry || 'index.html';
  const hasCname = opts.hasCname !== false;
  const found = [];
  const add = (id, file, message) => found.push({ id, file, message });
  const html = Object.entries(files).filter(([f]) => f.endsWith('.html'));
  const entry = files[entryName];

  // --- 秘密情報（v5 A2 / Part IV） -------------------------------------
  for (const [f, raw] of Object.entries(files)) {
    const src = stripComments(raw, f.endsWith('.html') ? 'html' : 'js');
    // Google の API キーは AIza で始まり 39 文字。末尾に \b を置くと
    // 長めのダミー文字列を取りこぼすので、始まりだけを見る
    if (/\bAIza[0-9A-Za-z_-]{35}/.test(src)) add('SECRET_APIKEY', f, 'API キーらしき文字列が直書きされている');
    // このルールが探しているのは「コードに紛れ込んだ連絡先」であって、
    // 法務ページに載せる問い合わせ先ではない。プライバシーポリシーと利用規約は
    // 連絡先を載せることが要件なので、対象から外す。
    const LEGAL_PAGES = ['privacy.html', 'terms.html'];
    // ドメイン側は英字で始まることを求める。そうしないと npm の版指定
    // （pkg@3.3.0 の "pkg@3.3.0" の部分）をメールと読み違える。
    // 繰り返しに上限を付けてある。上限なしだと、長い1行に当てたときに
    // 後戻りが爆発する（実測：400KB の1行で 98 秒かかっていた）。
    if (!LEGAL_PAGES.includes(f.split('/').pop())
      && /[\w.+-]{1,64}@(?!example\.)[\w-]{1,63}\.[A-Za-z][\w.]{0,63}/.test(src.replace(/https?:\/\/\S+/g, '')))
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
    // 残る外部資産（スタイル）は版の固定を要求する。
    // 見るのは rel="stylesheet" だけ。<link> にはほかに canonical・icon・
    // manifest などがあり、それらは「版を固定する」対象ではない。
    for (const m of src.matchAll(/<link\b[^>]*>/gi)) {
      const tag = m[0];
      if (!/\brel\s*=\s*["']?stylesheet\b/i.test(tag)) continue;
      const href = tag.match(/\bhref\s*=\s*["'](https?:\/\/[^"']+)["']/i);
      if (!href) continue;                                   // 自サイト内の相対パスは対象外
      const url = href[1].replace(/^https?:\/\//, '');
      if (/fonts\.googleapis\.com|fonts\.gstatic\.com/.test(url)) continue;  // フォントは見た目だけ
      if (!/@\d+\.\d+\.\d+|\/\d+\.\d+\.\d+\//.test(url))
        add('DEP_UNPINNED', f, '外部スタイルの版が固定されていない: ' + url.split('/')[0]);
    }
  }

  // --- 表示（v5 §2）：すべての HTML に課す ------------------------------
  for (const [f, raw] of html) {
    const src = stripComments(raw, 'html');
    if (/user-scalable\s*=\s*no|maximum-scale\s*=\s*1/.test(src))
      add('VIEWPORT_NOZOOM', f, '拡大を禁止している（見えづらい子が拡大できなくなる）');

    for (const m of src.matchAll(/<meta[^>]+name\s*=\s*["']viewport["'][^>]+content\s*=\s*["']([^"']+)["']/gi))
      if (!/viewport-fit\s*=\s*cover/.test(m[1])) add('VIEWPORT_NO_FIT', f, '<meta> の viewport に viewport-fit=cover が無い');

    // 100vh 単独は使わない。同じ行か次の行に dvh があればフォールバックとして正しい
    const lines = src.split('\n');
    lines.forEach((line, i) => {
      if (!/100vh/.test(line)) return;
      const near = line + '\n' + (lines[i + 1] || '') + '\n' + (lines[i - 1] || '');
      if (!/dvh/.test(near)) add('VIEWPORT_100VH', f, `100vh を単独で使っている（${i + 1}行目）`);
    });

    if (!/safe-area-inset/.test(src)) add('A11Y_NO_SAFE_AREA', f, 'safe-area-inset を使っていない（切り欠きに潜り込む）');
    if (!/prefers-reduced-motion/.test(src)) add('A11Y_NO_REDUCED_MOTION', f, 'prefers-reduced-motion に対応していない');
    if (!/touch-action/.test(src)) add('A11Y_NO_TOUCH_ACTION', f, 'touch-action を指定していない');
    if (!/clamp\(/.test(src)) add('A11Y_NO_FLUID_TYPE', f, 'clamp() による可変文字サイズが無い');
  }

  // --- 表示：アプリ本体にだけ課す ---------------------------------------
  if (entry !== undefined) {
    const src = stripComments(entry, 'html');
    if (!/min-height:\s*44px/.test(src)) add('A11Y_NO_TAP44', entryName, '44px 以上のタップ領域を指定していない');
    if (!/:focus-visible/.test(src)) add('A11Y_NO_FOCUS_VISIBLE', entryName, 'キーボード操作の焦点が見えない（:focus-visible が無い）');
    if (!/@media\s+print/.test(src)) add('NO_PRINT_CSS', entryName, '印刷用の CSS が無い（このアプリは紙に刷るもの）');
    if (!/visualViewport/.test(src)) add('NO_VISUAL_VIEWPORT', entryName, 'ソフトキーボードに追従していない（visualViewport）');
    if (!/Content-Security-Policy/i.test(src)) add('CSP_MISSING', entryName, 'Content-Security-Policy が入っていない');
    if (!/addEventListener\(\s*['"]pagehide['"]/.test(src))
      add('NO_PAGEHIDE_FLUSH', entryName, 'pagehide で保存を確定していない（Chromebook のタブ破棄で消える）');
    // 画面に出す canvas は端末の画素密度に合わせる。合わせないと粗く見える。
    if (/getContext\(\s*['"]2d['"]/.test(src) && !/devicePixelRatio/.test(src))
      add('CANVAS_NO_DPR', entryName, 'Canvas を使っているのに devicePixelRatio を見ていない');
  }

  // --- PWA（v5 §3） -----------------------------------------------------
  if (entry !== undefined) {
    const src = stripComments(entry, 'html');
    // インストールの合図は、外部スクリプトより前で捕まえる。
    // 後ろだと通信の遅い端末で取りこぼし、インストールボタンが出なくなる。
    const at = src.indexOf('beforeinstallprompt');
    const firstScript = src.search(/<script[^>]+src\s*=/i);
    if (at < 0) add('PWA_INSTALL_LATE', entryName, 'beforeinstallprompt を捕まえていない');
    else if (firstScript >= 0 && at > firstScript) add('PWA_INSTALL_LATE', entryName, 'beforeinstallprompt が外部スクリプトより後ろにある');

    if (!/pwa-installable/.test(src) || !/__deferredInstallPrompt/.test(src))
      add('PWA_NO_INSTALL_BUTTON', entryName, 'アプリ内にインストールボタンが無い');
    if (!/あたらしい\s*バージョン/.test(src))
      add('PWA_NO_UPDATE_NOTICE', entryName, '更新のお知らせを出していない（押されるまで入れ替えない形にすること）');
    if (!/SKIP_WAITING/.test(src))
      add('PWA_NO_SKIP_WAITING_UI', entryName, '画面側から SKIP_WAITING を送っていない');
    if (!/apple-touch-icon/.test(src))
      add('PWA_NO_APPLE_ICON', entryName, 'apple-touch-icon を指していない（iOS は maskable 非対応）');
  }

  const manifestRaw = files['manifest.webmanifest'];
  if (manifestRaw !== undefined) {
    let manifest = null;
    try { manifest = JSON.parse(manifestRaw); }
    catch (e) { add('PWA_MANIFEST_BROKEN', 'manifest.webmanifest', 'JSON として読めない: ' + e.message); }
    if (manifest) {
      // 正しい値は「どこで配信するか」で変わる。
      // CNAME があれば独自ドメインの直下なので "./"。
      // オリジンを他アプリと共有する配置なら、取り違えを避けるため
      // リポジトリ名の絶対パスが要る。
      // ⚠️ 独自ドメインでリポジトリ名の絶対パスに戻すと、scope がページの URL を
      //    含まなくなり、manifest ごと無視されて PWA として入れられなくなる。
      const want = hasCname ? './' : `/${opts.repoName || ''}/`;
      for (const key of ['id', 'scope', 'start_url']) {
        const v = manifest[key];
        if (typeof v !== 'string' || !v.startsWith(want))
          add('PWA_MANIFEST_PATH', 'manifest.webmanifest', `${key}=${v}（${want} で始まること）`);
      }
      const need = ['192x192 any', '512x512 any', '192x192 maskable', '512x512 maskable'];
      const got = (manifest.icons || []).map((i) => `${i.sizes} ${i.purpose || 'any'}`);
      const missing = need.filter((n) => !got.includes(n));
      if (missing.length) add('PWA_ICONS', 'manifest.webmanifest', 'アイコンが足りない: ' + missing.join(' / '));
    }
  }

  // --- Service Worker ---------------------------------------------------
  for (const [f, raw] of Object.entries(files)) {
    if (!/(^|\/)sw\.js$/.test(f)) continue;
    const src = stripComments(raw, 'js');
    // 「消す式」ではなく「startsWith で絞る式があるか」を見る（v5 §P4 の実例）
    if (/caches\.keys\(\)/.test(src) && !/startsWith/.test(src))
      add('SW_CACHE_WIPE', f, '自アプリ以外のキャッシュまで消している疑い（startsWith で絞ること）');
    if (/localStorage/.test(src)) add('SW_LOCALSTORAGE', f, 'Service Worker が localStorage に触れている');
    // 写真は IndexedDB に置いてある。キャッシュの掃除のついでにここへ手を出すと、
    // 子どもの記事の写真が消える。localStorage と同じ理由で触らせない。
    if (/indexedDB/.test(src)) add('SW_INDEXEDDB', f, 'Service Worker が IndexedDB に触れている');
    if (!/SKIP_WAITING/.test(src)) add('SW_NO_SKIP_WAITING', f, 'SKIP_WAITING を受け取れない（更新を適用できない）');
    if (!/offline\.html/.test(src)) add('SW_NO_OFFLINE_PAGE', f, 'offline.html を先読みしていない');
    if (!/APP_VERSION\s*=\s*['"][^'"]+['"]/.test(src)) add('SW_NO_APP_VERSION', f, 'APP_VERSION が無い');
  }

  // --- ふりがな（v5 §4） ------------------------------------------------
  for (const [f, raw] of html) {
    // CSS の規則として読むので、<style> の中だけを対象にする。
    // ファイル全体に当てると、JavaScript の中の HTML 文字列を規則と読み違える。
    const src = [...stripComments(raw, 'html').matchAll(/<style[^>]*>([\s\S]*?)<\/style>/gi)]
      .map((m) => m[1]).join('\n');
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

  // --- 禁止事項（Part IV） ---------------------------------------------
  for (const [f, raw] of Object.entries(files)) {
    const src = stripComments(raw, f.endsWith('.html') ? 'html' : 'js');
    if (/localStorage\.clear\(\)/.test(src)) add('BAN_LS_CLEAR', f, 'localStorage.clear() を使っている（他アプリの保存まで消す）');
    if (/postMessage\([^)]*,\s*['"]\*['"]\s*\)/.test(src)) add('BAN_POSTMESSAGE', f, 'postMessage の宛先が * になっている');
  }

  // --- 構文 -------------------------------------------------------------
  // HTML の中のインラインの <script> を見る（このアプリはコードを全部そこに持つ）。
  for (const [f, raw] of html) {
    // コメントを先に落とす。落とさないと、コメントの中に書いた <script> の
    // 文字を本物のタグと読み違え、そこからコメントの残りごと拾ってしまう。
    for (const m of stripComments(raw, 'html').matchAll(/<script(?![^>]*\bsrc\s*=)[^>]*>([\s\S]*?)<\/script>/gi)) {
      try { new Script(m[1]); }
      catch (e) { add('SYNTAX', f, 'インラインの <script> に構文エラー: ' + e.message); }
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
    ['.gitignore', '作業用のファイルが混入する'],
    ['README.md', '開発者向けの説明が無い'],
    ['MANUAL.md', '使う人向けの説明が無い'],
    ['AUDIT.md', '実測の記録が無い'],
    ['.github/dependabot.yml', '依存の更新が来ない'],
    ['manifest.webmanifest', 'PWA として入れられない'],
    ['sw.js', 'オフラインで開けない'],
    ['offline.html', '圏外で開いたときに白い画面になる'],
    ['icons/apple-touch-icon.png', 'iOS のホーム画面で絵が出ない'],
  ]) if (!exists(f)) missing.push({ id: 'FILE_MISSING', file: f, message: why });
  return missing;
}

/**
 * 画像の大きさ。校内 Wi-Fi は混む。1枚が重いと初回表示がそのぶん遅れる。
 * assets = [{ path, bytes }]
 */
export function checkAssets(assets) {
  const LIMITS = [
    [/^favicon\.png$/, 30 * 1024],
    [/^icons\/(icon|maskable)-512\.png$/, 60 * 1024],
    [/\.(png|jpe?g|gif|webp)$/i, 150 * 1024],
  ];
  const found = [];
  for (const { path, bytes } of assets) {
    for (const [re, max] of LIMITS) {
      if (!re.test(path)) continue;
      if (bytes > max) found.push({
        id: 'ASSET_TOO_BIG', file: path,
        message: `${Math.round(bytes / 1024)}KB（めやす ${Math.round(max / 1024)}KB を超えている）`,
      });
      break;
    }
  }
  return found;
}

const RULES = [
  'SECRET_APIKEY / SECRET_MAIL … 秘密情報の直書き',
  'DEP_BABEL / DEP_TAILWIND_CDN / DEP_CDN_SCRIPT / DEP_UNPINNED … 依存',
  'VIEWPORT_NOZOOM / VIEWPORT_NO_FIT / VIEWPORT_100VH … 表示',
  'A11Y_NO_SAFE_AREA / A11Y_NO_REDUCED_MOTION / A11Y_NO_TOUCH_ACTION / A11Y_NO_FLUID_TYPE',
  'A11Y_NO_TAP44 / A11Y_NO_FOCUS_VISIBLE / NO_PRINT_CSS / NO_VISUAL_VIEWPORT / NO_PAGEHIDE_FLUSH / CANVAS_NO_DPR',
  'CSP_MISSING … Content-Security-Policy',
  'PWA_INSTALL_LATE / PWA_NO_INSTALL_BUTTON / PWA_NO_UPDATE_NOTICE / PWA_NO_SKIP_WAITING_UI / PWA_NO_APPLE_ICON',
  'PWA_MANIFEST_BROKEN / PWA_MANIFEST_PATH / PWA_ICONS … manifest',
  'SW_CACHE_WIPE / SW_LOCALSTORAGE / SW_INDEXEDDB / SW_NO_SKIP_WAITING / SW_NO_OFFLINE_PAGE / SW_NO_APP_VERSION … Service Worker',
  'RUBY_HARDCODED … ふりがなの色',
  'BAN_LS_CLEAR / BAN_POSTMESSAGE … 禁止事項',
  'SYNTAX … インラインの <script> の構文',
  'SIZE_LINES / SIZE_BYTES / ASSET_TOO_BIG … 大きさ',
  'FILE_MISSING … 置かれているべきファイル',
];

// --- ここから下は実行用（テストからは import されない） -------------------
if (import.meta.url === `file://${process.argv[1]}`) {
  if (process.argv.includes('--list')) {
    console.log(RULES.join('\n'));
    process.exit(0);
  }

  const cfg = existsSync(join(ROOT, 'quality.config.json'))
    ? JSON.parse(readFileSync(join(ROOT, 'quality.config.json'), 'utf8'))
    : {};
  const waived = new Map((cfg.waived || []).map((w) => [w.rule, w.reason]));

  const files = {};
  const assets = [];
  const SKIP_DIRS = new Set(['.git', 'node_modules', 'scripts', 'tests', 'vendor', '.github']);
  const walk = (rel) => {
    for (const e of readdirSync(join(ROOT, rel || '.'), { withFileTypes: true })) {
      const r = rel ? `${rel}/${e.name}` : e.name;
      if (e.isDirectory()) {
        if (SKIP_DIRS.has(e.name)) continue;
        walk(r);
      } else if (/\.(html|js|mjs|json|webmanifest)$/.test(e.name)) {
        files[r] = readFileSync(join(ROOT, r), 'utf8');
      } else if (/\.(png|jpe?g|gif|webp|svg)$/i.test(e.name)) {
        assets.push({ path: r, bytes: statSync(join(ROOT, r)).size });
      }
    }
  };
  walk('');

  const found = [
    ...checkFiles((f) => existsSync(join(ROOT, f))),
    ...inspect(files, { entry: cfg.entry || 'index.html', hasCname: existsSync(join(ROOT, 'CNAME')), repoName: cfg.repoName }),
    ...checkAssets(assets),
  ];

  const live = found.filter((p) => !waived.has(p.id));
  const skipped = found.filter((p) => waived.has(p.id));

  console.log(`検査したファイル: ${Object.keys(files).length}本 / 画像 ${assets.length}点`);
  for (const p of live) console.log(`❌ [${p.id}] ${p.file}: ${p.message}`);
  if (skipped.length) {
    console.log('\n免除した項目（quality.config.json に理由を明記済み）:');
    for (const p of skipped) console.log(`  - [${p.id}] ${p.file}: ${p.message}\n    理由: ${waived.get(p.id)}`);
  }
  if (live.length === 0) {
    console.log('✅ 指摘なし');
    process.exit(0);
  }
  console.log(`\n${live.length}件`);
  process.exit(1);
}
