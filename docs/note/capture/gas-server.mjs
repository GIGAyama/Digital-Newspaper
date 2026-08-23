/* 撮影用のローカル・スタンドイン。
 *
 *   node docs/note/capture/gas-server.mjs 4180
 *
 * このリポジトリは Google Apps Script のウェブアプリで、画面（post.html /
 * admin.html）は `google.script.run` 越しに Code.gs を呼ぶ。ビルド成果物が
 * 無いので serve.mjs では配れない。そこで、
 *
 *   - doGet(e) の振り分け（/exec と /exec?p=admin）
 *   - Code.gs の各関数（getArticles / saveArticle / …）と同じ入出力
 *   - スプレッドシートとドライブの代わりのメモリ上の置き場
 *
 * だけを持つ小さなサーバを立て、画面そのものは手を入れずに配る。
 * 画面の HTML は、次の3点だけを書きかえて配る（それ以外は原文のまま）。
 *
 *   1. <?= ScriptApp.getService().getUrl() ?> を /exec に置きかえる（GAS の
 *      テンプレート展開の代わり）
 *   2. unpkg / Google Fonts の URL を、npm から取った同じ版のローカル写しに向ける
 *      （この環境から外部 CDN に出られないため。中身は同じもの）
 *   3. google.script.run の代わりに、このサーバへ問い合わせる同じ形の呼び出しを足す
 */
import http from 'node:http';
import { readFileSync, existsSync, statSync, readdirSync } from 'node:fs';
import { join, extname, resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(HERE, '../../..');   // リポジトリの根
const PORT = Number(process.argv[2] || 4180);

// ---------------------------------------------------------------- 置き場（スプレッドシート／ドライブの代わり）
const iso = (s) => new Date(s).getTime();
const uploads = new Map(); // name -> {mime, buf}

for (const f of readdirSync(join(HERE, 'assets'))) {
  uploads.set(f, { mime: 'image/svg+xml', buf: readFileSync(join(HERE, 'assets', f)) });
}

let seq = 0;
const uuid = () => `sample-${String(++seq).padStart(4, '0')}`;

/* Articles シートの中身。撮影のために用意した見本で、実在の学級のものではない。 */
const articles = [
  {
    id: uuid(), title: 'うんどう会 さいごまで 走りきったよ',
    body: '七月十日に、うんどう会がありました。ぼくは百メートル走に出ました。スタートのピストルが鳴ったとき、心ぞうがドキドキしました。とちゅうで足がいたくなったけれど、クラスのみんなが名前をよんでくれたので、さいごまで走りきることができました。ゴールしたあと、先生が「よくがんばったね」と言ってくれてうれしかったです。来年はもっと速く走れるように、休み時間にれん習しようと思います。',
    imageUrl: '/uploads/undokai.svg', reporterName: 'たけうち はると',
    timestamp: iso('2026-07-10T14:20:00+09:00'), tag: '行事',
  },
  {
    id: uuid(), title: '学級園の ミニトマトが 赤くなりました',
    body: '五月にうえたミニトマトが、やっと赤くなりました。毎朝、当番の人が水やりをしています。はじめは小さな白い花だったのに、だんだんみどりの実になって、いまは十二こも赤くなっています。理科の先生に聞いたら、「よく日に当てるとあまくなるよ」と教えてくれました。来週、みんなで食べてみるのが楽しみです。',
    imageUrl: '/uploads/tomato.svg', reporterName: 'もりた あおい',
    timestamp: iso('2026-07-13T09:05:00+09:00'), tag: '学習',
  },
  {
    id: uuid(), title: '休み時間の ドッジボールが 大にんき',
    body: '中休みになると、みんながグラウンドに走っていきます。いま三年二組ではドッジボールがはやっています。先週から男女まぜてチームを作るようにしたので、前よりもりあがるようになりました。ボールが当たっても「ドンマイ」と言い合えるのがいいところだと思います。あしたも晴れますように。',
    imageUrl: '/uploads/dodgeball.svg', reporterName: 'かとう りく',
    timestamp: iso('2026-07-14T12:40:00+09:00'), tag: '遊び',
  },
  {
    id: uuid(), title: 'なわとび記録会に むけて 練習中',
    body: '九月のなわとび記録会にむけて、体育の時間に二重とびのれん習をしています。わたしはまだ三回しかとべません。でも、上手な人に足のうごかし方を教えてもらってから、少しずつつづくようになりました。目ひょうは十回です。おうちでもれん習しています。',
    imageUrl: '/uploads/nawatobi.svg', reporterName: 'ふじた ひなの',
    timestamp: iso('2026-07-15T15:10:00+09:00'), tag: '学習',
  },
  {
    id: uuid(), title: '新しい ALTの先生が 来ました',
    body: '今週から、新しいALTの先生が来ました。名前はエマ先生です。オーストラリアから来たそうです。じこしょうかいのとき、すきな食べものは「おにぎり」と言っていて、みんなでわらいました。英語の時間がもっと楽しみになりました。',
    imageUrl: '', reporterName: 'すずき みなと',
    timestamp: iso('2026-07-15T16:30:00+09:00'), tag: 'ニュース',
  },
  {
    id: uuid(), title: 'そうじの時間 ゆかピカピカ大作せん',
    body: 'そうじの時間に、教室のゆかを新聞紙でみがいてみました。用むいんさんに教えてもらったやり方です。水にぬらした新聞紙をちぎってまくと、ほこりがくっついて取れやすくなります。やってみたら、いつもよりずっときれいになりました。来週も続けたいです。',
    imageUrl: '', reporterName: 'いのうえ そうた',
    timestamp: iso('2026-07-16T13:55:00+09:00'), tag: 'その他',
  },
  {
    id: uuid(), title: '図書室の本を 百さつ 読みました',
    body: 'この一学期で、図書室の本を百さつ読みました。読書カードがぜんぶうまったとき、司書の先生がシールをくれました。いちばんおもしろかったのは、こん虫の図かんです。カブトムシの角が、じつは頭ではなく口のちかくから出ていることを知っておどろきました。二学期は百五十さつを目ひょうにします。',
    imageUrl: '', reporterName: 'おおた ゆい',
    timestamp: iso('2026-07-16T14:15:00+09:00'), tag: '学習',
  },
];

const DEFAULT_TAGS = [
  { icon: '📰', name: 'ニュース', ruby: 'ニュース' },
  { icon: '🎌', name: '行事', ruby: 'ぎょうじ' },
  { icon: '✏️', name: '学習', ruby: 'がくしゅう' },
  { icon: '⚽', name: '遊び', ruby: 'あそび' },
  { icon: '🍀', name: 'その他', ruby: 'そのた' },
];

let tagSettings = null;                 // スクリプトプロパティの代わり
const systemRows = [];                  // SystemData シートの代わり

const fmt = (d) => {
  const p = (n) => String(n).padStart(2, '0');
  return `${p(d.getMonth() + 1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
};

// ---------------------------------------------------------------- Code.gs と同じ入出力
const API = {
  getTagsSettings: () => (tagSettings ? JSON.parse(tagSettings) : DEFAULT_TAGS),
  saveTagsSettings: (json) => { tagSettings = json; return { success: true }; },

  saveArticle: (data) => {
    const id = uuid();
    let imageUrl = '';
    if (data.image) {
      const name = `${id}.${(data.mimeType || 'image/png').split('/')[1].replace('+xml', '')}`;
      uploads.set(name, { mime: data.mimeType, buf: Buffer.from(data.image, 'base64') });
      imageUrl = `/uploads/${name}`;
    }
    articles.push({
      id, title: data.title, body: data.body, imageUrl,
      reporterName: data.reporter, timestamp: Date.now(), tag: data.tag || '',
    });
    return { success: true };
  },

  getArticles: () => [...articles].reverse().map((a) => ({ ...a })),

  updateArticleTag: (id, newTag) => {
    const a = articles.find((x) => x.id === id);
    if (a) a.tag = newTag;
    return null;
  },

  saveLayoutState: (name, json) => {
    systemRows.push(['LAYOUT', name, json, new Date()]);
    return { message: '✅ 保存しました' };
  },
  getSavedList: () => systemRows.filter((r) => r[0] === 'LAYOUT')
    .map((r) => ({ name: r[1], date: fmt(r[3]) })).reverse(),
  loadLayoutState: (name) => {
    for (let i = systemRows.length - 1; i >= 0; i--) {
      if (systemRows[i][0] === 'LAYOUT' && systemRows[i][1] === name) return { success: true, data: systemRows[i][2] };
    }
    return { success: false, message: 'データが見つかりません' };
  },
  saveTemplate: (name, json) => {
    for (let i = systemRows.length - 1; i >= 0; i--) {
      if (systemRows[i][0] === 'TEMPLATE' && systemRows[i][1] === name) systemRows.splice(i, 1);
    }
    systemRows.push(['TEMPLATE', name, json, new Date()]);
    return { message: '✅ テンプレートを保存しました' };
  },
  getTemplateList: () => systemRows.filter((r) => r[0] === 'TEMPLATE').map((r) => ({ name: r[1] })).reverse(),
  loadTemplate: (name) => {
    for (let i = systemRows.length - 1; i >= 0; i--) {
      if (systemRows[i][0] === 'TEMPLATE' && systemRows[i][1] === name) return { success: true, data: systemRows[i][2] };
    }
    return { success: false, message: 'テンプレートが見つかりません' };
  },
};

// ---------------------------------------------------------------- 画面を配る
const GAS_SHIM = `<script>
/* google.script.run の代わり。呼び出しの形（withSuccessHandler(...).method(...)）
   と、返ってくる値の形は Code.gs と同じ。 */
(function () {
  var METHODS = ${JSON.stringify(Object.keys(API))};
  function makeRunner() {
    var ok = null, fail = null;
    var r = {
      withSuccessHandler: function (f) { ok = f; return r; },
      withFailureHandler: function (f) { fail = f; return r; }
    };
    METHODS.forEach(function (m) {
      r[m] = function () {
        var args = Array.prototype.slice.call(arguments);
        fetch('/__gas/' + m, {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(args)
        }).then(function (res) { return res.json(); })
          .then(function (res) {
            if (res.error) { if (fail) fail(new Error(res.error)); }
            else if (ok) ok(res.result);
          })
          .catch(function (e) { if (fail) fail(e); });
      };
    });
    return r;
  }
  window.google = { script: {} };
  Object.defineProperty(window.google.script, 'run', { get: makeRunner });
})();
</script>`;

const FONTS_POST = ['biz-udpgothic/400', 'biz-udpgothic/700'];
const FONTS_ADMIN = [
  'shippori-mincho/400', 'shippori-mincho/700',
  'noto-sans-jp/400', 'noto-sans-jp/700',
  'yomogi/400', 'kiwi-maru/400', 'kiwi-maru/500',
  'biz-udpgothic/400', 'biz-udpgothic/700',
];
const fontCss = (list) => list.map((p) => `@import url("/vendor/f/${p}.css");`).join('\n');

const renderPage = (name) => {
  let html = readFileSync(join(REPO, `${name}.html`), 'utf8');
  html = html.replace(/<\?=\s*ScriptApp\.getService\(\)\.getUrl\(\)\s*\?>/g, '/exec');
  html = html.replace('https://unpkg.com/@picocss/pico@1.5.10/css/pico.min.css', '/vendor/pico.min.css');
  html = html.replace(/https:\/\/fonts\.googleapis\.com\/css2\?[^"']*/g,
    name === 'admin' ? '/vendor/fonts-admin.css' : '/vendor/fonts-post.css');
  html = html.replace('<meta charset="UTF-8">', `<meta charset="UTF-8">\n${GAS_SHIM}`);
  return html;
};

const MIME = {
  '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8', '.json': 'application/json; charset=utf-8',
  '.woff2': 'font/woff2', '.woff': 'font/woff', '.svg': 'image/svg+xml',
  '.png': 'image/png', '.jpg': 'image/jpeg', '.map': 'application/json',
};

const send = (res, code, type, body) => {
  res.writeHead(code, { 'Content-Type': type, 'Content-Length': Buffer.byteLength(body), 'Cache-Control': 'no-store' });
  res.end(body);
};

const server = http.createServer((req, res) => {
  const url = new URL(req.url, 'http://127.0.0.1');
  const path = decodeURIComponent(url.pathname);

  if (req.method === 'POST' && path.startsWith('/__gas/')) {
    const fn = API[path.slice('/__gas/'.length)];
    let body = '';
    req.on('data', (c) => { body += c; });
    req.on('end', () => {
      if (!fn) return send(res, 404, MIME['.json'], JSON.stringify({ error: 'no such function' }));
      /* GAS の往復はだいたい 1〜3 秒かかる。すぐ返してしまうと
         「⏳ 送信中...」や「通信中...」が画面に出ている時間が無くなり、
         本番では見えるはずの状態が撮れない。待ちを入れて近づける。 */
      const delay = path.endsWith('/saveArticle') ? 900 : 250;
      setTimeout(() => {
        try {
          const result = fn(...JSON.parse(body || '[]'));
          send(res, 200, MIME['.json'], JSON.stringify({ result: result ?? null }));
        } catch (e) {
          send(res, 200, MIME['.json'], JSON.stringify({ error: String(e) }));
        }
      }, delay);
    });
    return;
  }

  if (path === '/' || path === '/exec') {
    return send(res, 200, MIME['.html'], renderPage(url.searchParams.get('p') === 'admin' ? 'admin' : 'post'));
  }
  if (path === '/vendor/fonts-post.css') return send(res, 200, MIME['.css'], fontCss(FONTS_POST));
  if (path === '/vendor/fonts-admin.css') return send(res, 200, MIME['.css'], fontCss(FONTS_ADMIN));
  if (path === '/vendor/pico.min.css') {
    return send(res, 200, MIME['.css'], readFileSync(join(REPO, 'node_modules/@picocss/pico/css/pico.min.css')));
  }
  if (path.startsWith('/vendor/f/')) {
    const file = join(REPO, 'node_modules/@fontsource', path.slice('/vendor/f/'.length));
    if (existsSync(file) && statSync(file).isFile()) return send(res, 200, MIME[extname(file)] || 'application/octet-stream', readFileSync(file));
  }
  if (path.startsWith('/uploads/')) {
    const f = uploads.get(path.slice('/uploads/'.length));
    if (f) return send(res, 200, f.mime, f.buf);
  }
  send(res, 404, 'text/plain; charset=utf-8', 'not found');
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`記者ページ  http://127.0.0.1:${PORT}/exec`);
  console.log(`編集室      http://127.0.0.1:${PORT}/exec?p=admin`);
});
