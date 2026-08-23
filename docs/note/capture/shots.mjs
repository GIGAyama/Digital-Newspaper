/* 紹介記事に載せる画面を撮る。
 * 画面はすべて本物。フォームに打ちこみ、ボタンを押して撮っている。
 * 差し替えたのは (1) Google Fonts をローカル写しに (2) PeerJS の接続先を
 * 手元の PeerServer に、の2点だけで、アプリのコードには手を入れていない。
 */
import { chromium } from 'playwright';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';

const BASE = process.env.CAPTURE_BASE || 'http://127.0.0.1:8098/';
const OUT = path.resolve(process.argv[2] || 'docs/note/images');
fs.mkdirSync(OUT, { recursive: true });

const peerShim = () => {
  let Native = null, shim = null;
  Object.defineProperty(window, 'Peer', {
    configurable: true,
    get() { return shim; },
    set(v) {
      Native = v;
      shim = function (id, opts) {
        const o = Object.assign({ host: '127.0.0.1', port: 9000, path: '/peer', secure: false, debug: 0 }, opts || {});
        return id === undefined ? new Native(o) : new Native(id, o);
      };
    },
  });
};

/* 見本の写真。児童の顔が写った本物は使えないので、絵で用意する。 */
const DRAW = {
  undokai: `const g=x.createLinearGradient(0,0,0,H);g.addColorStop(0,'#bfe3f7');g.addColorStop(1,'#cfe8c2');x.fillStyle=g;x.fillRect(0,0,W,H);
    x.fillStyle='#d9a066';x.beginPath();x.ellipse(W/2,H*0.72,W*0.46,H*0.30,0,0,7);x.fill();
    x.strokeStyle='#fff';x.lineWidth=6;x.beginPath();x.ellipse(W/2,H*0.72,W*0.34,H*0.20,0,0,7);x.stroke();
    x.fillStyle='#e74c3c';x.beginPath();x.moveTo(W*0.14,H*0.28);x.lineTo(W*0.14,H*0.62);x.lineTo(W*0.10,H*0.62);x.lineTo(W*0.10,H*0.28);x.fill();
    x.beginPath();x.moveTo(W*0.14,H*0.28);x.lineTo(W*0.30,H*0.34);x.lineTo(W*0.14,H*0.42);x.fill();
    x.fillStyle='#fff';x.beginPath();x.arc(W*0.78,H*0.20,H*0.10,0,7);x.fill();`,
  tomato: `x.fillStyle='#f3f7ee';x.fillRect(0,0,W,H);
    x.strokeStyle='#7cb342';x.lineWidth=8;x.beginPath();x.moveTo(W*0.5,H);x.quadraticCurveTo(W*0.44,H*0.5,W*0.5,H*0.18);x.stroke();
    [[0.36,0.62,52],[0.64,0.55,46],[0.5,0.34,40]].forEach(function(p,i){
      x.fillStyle=i===2?'#ef5350':'#e53935';x.beginPath();x.arc(W*p[0],H*p[1],p[2],0,7);x.fill();
      x.fillStyle='#43a047';x.beginPath();x.arc(W*p[0],H*p[1]-p[2]*0.9,p[2]*0.35,0,7);x.fill();});
    x.fillStyle='#66bb6a';[[0.3,0.42],[0.7,0.36]].forEach(function(p){x.beginPath();x.ellipse(W*p[0],H*p[1],46,20,0.4,0,7);x.fill();});`,
  pool: `const g=x.createLinearGradient(0,0,0,H);g.addColorStop(0,'#e3f2fd');g.addColorStop(0.35,'#b3e5fc');g.addColorStop(1,'#0288d1');x.fillStyle=g;x.fillRect(0,0,W,H);
    x.strokeStyle='rgba(255,255,255,.7)';x.lineWidth=7;
    for(let i=1;i<6;i++){x.beginPath();for(let t=0;t<=W;t+=10){x.lineTo(t,H*0.4+i*H*0.11+Math.sin(t/38+i)*7);}x.stroke();}
    x.fillStyle='#ffca28';x.beginPath();x.arc(W*0.82,H*0.16,H*0.11,0,7);x.fill();`,
};

const photoOf = (kind) => `(function(){var c=document.createElement('canvas');c.width=1200;c.height=900;
  var x=c.getContext('2d');var W=1200,H=900;${DRAW[kind]}return c.toDataURL('image/png');})()`;

const browser = await chromium.launch();

const newDevice = async (label, vp) => {
  const ctx = await browser.newContext({ viewport: vp || { width: 1200, height: 900 }, locale: 'ja-JP', deviceScaleFactor: 1 });
  await ctx.addInitScript(peerShim);
  const page = await ctx.newPage();
  page.on('pageerror', (e) => console.log(`[${label}] pageerror`, e.message));
  await page.goto(BASE, { waitUntil: 'networkidle' });
  await page.waitForTimeout(700);
  await page.evaluate(() => document.fonts.ready);
  return { ctx, page };
};

const shot = async (page, name, opts = {}) => {
  await page.waitForTimeout(opts.wait || 400);
  // 合図（トースト）が消えるまで待つ。3秒ちょっとで自分から消える。
  // 待たずに撮ると、紙面の上に緑の帯が乗ったままの絵になる。
  await page.waitForFunction(() => {
    const c = document.getElementById('toastContainer');
    return !c || c.children.length === 0;
  }, { timeout: 8000 }).catch(() => {});
  const file = path.join(OUT, name + '.png');
  if (opts.selector) await page.locator(opts.selector).screenshot({ path: file });
  else await page.screenshot({ path: file, fullPage: !!opts.full, clip: opts.clip });
  const kb = Math.round(fs.statSync(file).size / 1024);
  console.log(`  ${name}.png  ${kb}KB${kb > 150 ? '  ← 大きすぎる' : ''}`);
};

const write = async (page, a) => {
  await page.click('#tabWrite'); await page.waitForTimeout(200);
  await page.click('#btnNew'); await page.waitForTimeout(200);
  await page.fill('#fTitle', a.title);
  await page.fill('#fBody', a.body);
  await page.fill('#fReporter', a.reporter);
  await page.click(`#tagGrid label:nth-of-type(${a.tag})`);
  if (a.photo) {
    const dataUrl = await page.evaluate(photoOf(a.photo));
    const tmp = path.join(os.tmpdir(), `dnp-shot-${a.photo}.png`);
    fs.writeFileSync(tmp, Buffer.from(dataUrl.split(',')[1], 'base64'));
    await page.setInputFiles('#fPhoto', tmp);
    await page.waitForFunction(() => !document.getElementById('photoPreview').hidden, { timeout: 8000 });
  }
  await page.click('#btnSave');
  await page.waitForTimeout(350);
};

// ===== あつめる側（紙面を組む端末） =====
const host = await newDevice('host', { width: 1400, height: 1000 });

// 01 まっさらな「記事を かく」
await host.page.setViewportSize({ width: 420, height: 1500 });
await shot(host.page, '01-write-empty', { wait: 700 });

// 02 記入したところ
await host.page.fill('#fTitle', '運動会で リレーの アンカーを つとめたよ');
await host.page.fill('#fBody', 'きのうの 運動会で、赤組の アンカーを つとめました。\nバトンを もらったときは 三位でしたが、さいごの コーナーで 二人 ぬいて 一位に なれました。\n毎朝の 練習の せいかが 出て、うれしかったです。');
await host.page.fill('#fReporter', 'やまだ はなこ');
await host.page.click('#tagGrid label:nth-of-type(2)');
{
  const dataUrl = await host.page.evaluate(photoOf('undokai'));
  const tmp = path.join(os.tmpdir(), 'dnp-shot-undokai.png');
  fs.writeFileSync(tmp, Buffer.from(dataUrl.split(',')[1], 'base64'));
  await host.page.setInputFiles('#fPhoto', tmp);
  await host.page.waitForFunction(() => !document.getElementById('photoPreview').hidden, { timeout: 8000 });
}
await shot(host.page, '02-write-filled', { wait: 600 });
await host.page.click('#btnSave');
await host.page.waitForTimeout(400);

await host.page.setViewportSize({ width: 1400, height: 1000 });
for (const a of [
  { title: 'ミニトマトが 赤く なりました', body: '生活科で そだてている ミニトマトが、やっと 赤く なりました。\n毎日 水を あげた かいが ありました。あまくて おいしかったです。', reporter: 'さとう たろう', tag: 3, photo: 'tomato' },
  { title: '大なわとびの 記録が 百回を こえた', body: '中休みに みんなで 大なわとびを しています。\nきのう はじめて 百回を こえました。かけ声を そろえたのが よかったと 思います。', reporter: 'いのうえ みなと', tag: 4 },
  { title: '図書室に 新しい 本が 入りました', body: '図書室に 新しい 本が 五十さつ 入りました。\nこん虫の 図かんが 人気で、休み時間は じゅんばん待ちです。', reporter: 'なかむら あおい', tag: 1 },
]) await write(host.page, a);

// 03 わたしの記事
await host.page.setViewportSize({ width: 620, height: 1200 });
await host.page.evaluate(() => document.getElementById('mineList').scrollIntoView({ block: 'center' }));
await shot(host.page, '03-mine-list', { wait: 600, selector: '#viewWrite .card:nth-of-type(2)' });

// 04 へやをひらく
await host.page.setViewportSize({ width: 980, height: 1000 });
await host.page.click('#btnConnect');
await host.page.click('#btnHost');
await host.page.waitForSelector('#hostActive:not([hidden])', { timeout: 25000 });
const code = (await host.page.textContent('#roomCodeText')).trim();
console.log('  あいことば:', code);
await shot(host.page, '04-room-open', { wait: 700, selector: '#connectModal .modal-box' });

// ===== おくる側 =====
const guest = await newDevice('guest', { width: 420, height: 1400 });
await write(guest.page, { title: 'プール開き。水が つめたかった', body: 'きょうは プール開きでした。\n入った しゅんかんは とても つめたかったけれど、なれると 気もちよかったです。\nことしは 二十五メートル 泳げるように なりたいです。', reporter: 'きむら りく', tag: 2, photo: 'pool' });
await write(guest.page, { title: 'そうじの 時間に 見つけた こと', body: 'ろうかを ふいていたら、去年の 学級文庫の 本が 出てきました。\nみんなで もとの たなに もどしました。おちている ものには 名前が 書いてあります。', reporter: 'たなか けん', tag: 5 });

// 05 あいことばを入れる
await guest.page.click('#btnConnect');
await guest.page.fill('#roomCodeInput', code.toLowerCase().replace('-', ''));
await shot(guest.page, '05-join', { wait: 500, selector: '#connectModal .modal-box' });
await guest.page.click('#btnGuest');
await guest.page.waitForSelector('#guestActive:not([hidden])', { timeout: 30000 });
await guest.page.waitForTimeout(1200);
await guest.page.click('[data-close="connectModal"]');
await guest.page.waitForTimeout(400);

// 記事を送る
await guest.page.evaluate(() => {
  document.querySelectorAll('#mineList .mine-item').forEach((it) => {
    const b = [...it.querySelectorAll('button')].find((e) => e.textContent.includes('おくる'));
    if (b) b.click();
  });
});
await guest.page.waitForTimeout(2500);

// 06 送ったあと
await guest.page.setViewportSize({ width: 620, height: 1000 });
await shot(guest.page, '06-sent', { wait: 800, selector: '#viewWrite .card:nth-of-type(2)' });

// 07 届いた記事（ホスト）
await host.page.click('[data-close="connectModal"]');
await host.page.waitForTimeout(400);
await host.page.setViewportSize({ width: 720, height: 1200 });
await host.page.click('#tabDesk');
await host.page.waitForTimeout(700);
await shot(host.page, '07-received', { wait: 700, selector: '#articleList' });

// 届いた2本を紙面にのせる
// 一覧はチェックのたびに組み直される。まとめて click() すると
// 2つめは外れた DOM を押すことになるので、1つずつ入れる。
for (let i = 0; i < 4; i++) {
  const done = await host.page.evaluate(() => {
    const box = document.querySelector('#articleList .article-item.from-peer input[type=checkbox]:not(:checked)');
    if (!box) return true;
    box.click();
    return false;
  });
  await host.page.waitForTimeout(900);
  if (done) break;
}
await host.page.waitForTimeout(1200);

// 08 のせてもらった（ゲスト）
await shot(guest.page, '08-published', { wait: 1200, selector: '#viewWrite .card:nth-of-type(2)' });

// ===== 紙面を組む =====
await host.page.setViewportSize({ width: 1400, height: 1250 });
await host.page.waitForTimeout(600);
await shot(host.page, '09-desk-panel', { wait: 800, selector: '.desk-panel' });

// 新聞名を打ちかえる
await host.page.evaluate(() => {
  const t = document.getElementById('paperTitle'); t.textContent = '三年二組 学級新聞'; t.dispatchEvent(new Event('blur'));
  const d = document.getElementById('paperDate'); d.textContent = '2026年 7月17日 一学期号'; d.dispatchEvent(new Event('blur'));
});
await host.page.waitForTimeout(500);

const paperShot = async (name) => {
  await host.page.setViewportSize({ width: 1000, height: 1300 });
  await host.page.waitForTimeout(700);
  await shot(host.page, name, { wait: 700, selector: '#paperArea' });
};
await paperShot('10-paper');

// 写真の置き方を変える
await host.page.evaluate(() => {
  const sels = document.querySelectorAll('#articleList .item-settings select');
  if (sels[0]) { sels[0].value = 'float-right'; sels[0].dispatchEvent(new Event('change')); }
});
await host.page.waitForTimeout(600);
await paperShot('11-paper-photos');

// QR を入れる
await host.page.setViewportSize({ width: 1400, height: 1250 });
await host.page.waitForTimeout(400);
await host.page.evaluate(() => {
  const ins = [...document.querySelectorAll('#articleList input[aria-label="QRコードにする URL"]')];
  const hit = ins[2] || ins[0];
  hit.value = 'https://digital-newspaper.giga-school.com/';
  hit.dispatchEvent(new Event('change'));
});
await host.page.waitForTimeout(800);
await paperShot('12-paper-qr');

// 黒板テーマ
await host.page.setViewportSize({ width: 1400, height: 1250 });
await host.page.selectOption('#themeSelect', 'blackboard');
await host.page.selectOption('#fontFamily', 'kiwi');
await host.page.waitForTimeout(600);
await paperShot('13-paper-blackboard');

// 春テーマ・横書き2段
await host.page.setViewportSize({ width: 1400, height: 1250 });
await host.page.selectOption('#themeSelect', 'spring');
await host.page.selectOption('#fontFamily', 'ud');
await host.page.selectOption('#textDir', 'horizontal');
await host.page.selectOption('#colCount', '2');
await host.page.waitForTimeout(600);
await paperShot('14-paper-spring-horizontal');

// もどす
await host.page.setViewportSize({ width: 1400, height: 1250 });
await host.page.selectOption('#themeSelect', 'default');
await host.page.selectOption('#fontFamily', 'mincho');
await host.page.selectOption('#textDir', 'vertical');
await host.page.selectOption('#colCount', '3');
await host.page.waitForTimeout(500);

// 15 自由記述らん
await host.page.setViewportSize({ width: 980, height: 900 });
await host.page.click('#btnFree');
await host.page.waitForTimeout(400);
await host.page.fill('#freeBody', '一学期の 新聞は これで さいごです。\nよんでくれて ありがとう。二学期も 楽しみに していてください。');
await shot(host.page, '15-free-modal', { wait: 500, selector: '#freeModal .modal-box' });
await host.page.click('#btnFreeAdd');
await host.page.waitForTimeout(600);
await paperShot('16-paper-free');

// 17 携帯の幅
await host.page.setViewportSize({ width: 390, height: 1400 });
await host.page.waitForTimeout(700);
await shot(host.page, '17-mobile-desk', { wait: 700 });

await browser.close();
console.log('DONE');
