/* note の紹介記事に載せる画面を撮るシナリオ。
 *
 *   node <skill>/scripts/capture.mjs docs/note/capture/shots.mjs \
 *     --base http://127.0.0.1:4180/exec --out docs/note/images
 *
 * 児童が1本投稿し、そのあと先生がその投稿を含めて紙面を組む、という順で通す。
 */
import { fileURLToPath } from 'node:url';

const asset = (name) => fileURLToPath(new URL(`assets/${name}`, import.meta.url));

export const viewport = { width: 390, height: 940 };

// 紙面だけを1枚に収める。A4 縦は 210mm×297mm ＝ おおよそ 794×1123px。
const paperShot = async (p, name) => {
  await p.resize(1000, 1260);
  await p.eval(() => {
    document.getElementById('paper-area').scrollIntoView({ block: 'start', behavior: 'instant' });
    window.scrollBy(0, -18);
  });
  await p.sleep(700);
  await p.shot(name);
};

// 記事一覧の列だけを見せたいとき。狭い画面では列が縦に積まれるので、
// その列を画面の上に寄せて撮る。
const listShot = async (p, name, inner = null) => {
  await p.resize(620, 1200);
  await p.eval(() => {
    document.getElementById('articleList').closest('.config-col').scrollIntoView({ block: 'start', behavior: 'instant' });
    window.scrollBy(0, -12);
  });
  if (inner) {
    await p.eval((needle) => {
      const box = document.getElementById('articleList');
      const item = [...box.querySelectorAll('.article-item')].find((e) => e.innerText.includes(needle));
      if (item) box.scrollTop = item.offsetTop - box.offsetTop - 6;
    }, inner);
  }
  await p.sleep(600);
  await p.shot(name);
};

const setQr = (p, title, url) => p.eval(([t, u]) => {
  const item = [...document.querySelectorAll('#articleList .article-item')].find((e) => e.innerText.includes(t));
  const el = item.querySelector('.qr-input');
  el.value = u;
  el.dispatchEvent(new Event('input', { bubbles: true }));
  return true;
}, [title, url]);

// 記事ごとの「配置／大／形」。0=配置 1=大 2=形
const setItemSelect = (p, title, idx, value) => p.eval(([t, i, v]) => {
  const item = [...document.querySelectorAll('#articleList .article-item')].find((e) => e.innerText.includes(t));
  const sel = item.querySelectorAll('.item-settings select')[i];
  if (!sel) return false;
  sel.value = v;
  sel.dispatchEvent(new Event('change', { bubbles: true }));
  return true;
}, [title, idx, value]);

const retype = async (p, selector, text) => {
  await p.raw.click(selector);
  await p.raw.keyboard.press('Control+A');
  await p.raw.keyboard.type(text);
};

export default async ({ open, log, base }) => {
  // ============================================================ 児童の画面
  const c = await open('児童');

  await c.resize(390, 1800);
  await c.shot('01-child-form');

  await c.raw.fill('#title', 'プール開きで 大きなプールに 入りました');
  await c.raw.fill('#body',
    '七月一日に、プール開きがありました。ことしから、となりの学校の大きなプールを'
    + 'かりることになりました。水がとてもつめたくて、はじめは足だけつけてがまんして'
    + 'いました。なれてきたら、けのびで五メートルすすむことができました。先生に'
    + '「バタ足のとき、ひざをのばすといいよ」と教えてもらったので、つぎの水泳の時間に'
    + 'ためしてみます。');
  await c.raw.fill('#reporter', 'なかむら ひかり');
  await c.click('行事');
  await c.raw.setInputFiles('#image', asset('pool.svg'));
  await c.sleep(600);
  await c.eval(() => { document.getElementById('body').scrollTop = 0; });
  await c.sleep(200);
  await c.shot('02-child-filled');

  await c.resize(390, 940);
  await c.scrollTo('送信する');
  await c.sleep(300);
  await c.click('送信する');
  await c.sleep(180);
  await c.shot('03-child-sending');
  await c.sleep(1400);
  await c.shot('04-child-success');
  log('送信後の児童画面:', (await c.text(200)).replace(/\n/g, ' '));

  // ============================================================ 先生の画面
  const t = await open('先生', { width: 1440, height: 1250, url: `${base}?p=admin` });
  log('編集室の記事:', await t.eval(() => [...document.querySelectorAll('#articleList .item-main label')].map((l) => l.innerText.trim())));

  await t.shot('05-admin-panel');
  await listShot(t, '06-admin-list');

  await t.resize(1440, 1250);
  await paperShot(t, '07-paper-initial');

  // 載せる記事を足す。写真のある3本を選ぶ
  await t.resize(1440, 1250);
  for (const title of ['うんどう会', '学級園の ミニトマト', 'なわとび記録会']) {
    const label = await t.eval((needle) => {
      const l = [...document.querySelectorAll('#articleList .item-main label')].find((e) => e.innerText.includes(needle));
      return l ? l.innerText.replace(/\s+/g, '') : null;
    }, title);
    if (!label) { log('見つからない:', title); continue; }
    await t.click(label);
    await t.sleep(500);
  }
  log('チェック数:', await t.eval(() => document.querySelectorAll('#articleList input[type=checkbox]:checked').length));

  // 新聞名と発行日を、その場で書きかえる
  await t.resize(1000, 1260);
  await t.eval(() => document.getElementById('paper-area').scrollIntoView({ block: 'start', behavior: 'instant' }));
  await t.sleep(300);
  await retype(t, '.newspaper-title', '三年二組 学級新聞');
  await retype(t, '.newspaper-date', '2026年 7月17日 一学期号');
  await t.raw.click('body');
  await t.sleep(500);
  await paperShot(t, '08-paper-renamed');

  // 写真の置きかたを変える
  await t.resize(1440, 1250);
  await setItemSelect(t, 'うんどう会', 0, 'float-right');
  await t.sleep(400);
  await setItemSelect(t, '学級園の ミニトマト', 1, '0.7');
  await t.sleep(400);
  await setItemSelect(t, 'なわとび記録会', 2, '1/1');
  await t.sleep(600);
  await listShot(t, '09-admin-imgsettings', 'うんどう会');
  await t.resize(1440, 1250);
  await paperShot(t, '10-paper-photos');

  // QR
  await t.resize(1440, 1250);
  await setQr(t, 'なわとび記録会', 'https://drive.google.com/file/d/1AbCdEfGhIjKlMnOpQrStU/view');
  await t.sleep(900);
  await listShot(t, '11-admin-qr-input', 'なわとび記録会');
  await t.resize(1440, 1250);
  await paperShot(t, '12-paper-qr');
  log('QR の塗られた画素:', await t.eval(() => {
    const cv = document.querySelector('.qr-container canvas');
    if (!cv) return 'canvas なし';
    const d = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height).data;
    let dark = 0;
    for (let i = 0; i < d.length; i += 4) if (d[i] < 128) dark++;
    return dark;
  }));

  // 絞りこみ
  await t.resize(620, 1200);
  await t.raw.fill('#searchInput', 'ドッジ');
  await t.sleep(700);
  await listShot(t, '13-admin-search');
  await t.raw.fill('#searchInput', '');
  await t.sleep(600);
  await t.raw.selectOption('#tagFilter', '学習');
  await t.sleep(700);
  await listShot(t, '14-admin-tagfilter');
  await t.raw.selectOption('#tagFilter', '');
  await t.sleep(600);

  // 自由記述欄
  await t.resize(1440, 1250);
  await t.eval(() => window.scrollTo(0, 0));
  await t.click('自由記述欄を追加');
  await t.sleep(800);
  await t.raw.fill('#freeBody',
    '一学期のあいだ、たくさんの記事がとどきました。読んでいると、教室のそとでも'
    + 'いろいろなことがあったのだと分かります。二学期も、みなさんの見つけたことを'
    + '待っています。');
  await t.sleep(400);
  await t.shot('15-admin-free-modal');
  await t.click('追加');
  await t.sleep(900);
  await paperShot(t, '16-paper-free');

  // レイアウトを変える
  await t.resize(1440, 1250);
  await t.raw.selectOption('#textDir', 'horizontal');
  await t.raw.selectOption('#colCount', '2');
  await t.sleep(800);
  await paperShot(t, '17-paper-horizontal');

  await t.resize(1440, 1250);
  await t.raw.selectOption('#textDir', 'vertical');
  await t.raw.selectOption('#colCount', '3');
  await t.sleep(700);

  // テーマ
  await t.raw.selectOption('#themeSelect', 'spring');
  await t.sleep(700);
  await paperShot(t, '18-paper-spring');
  await t.resize(1440, 1250);
  await t.raw.selectOption('#themeSelect', 'blackboard');
  await t.raw.selectOption('#fontFamily', 'kiwi');
  await t.sleep(800);
  await paperShot(t, '19-paper-blackboard');
  await t.resize(1440, 1250);
  await t.raw.selectOption('#themeSelect', 'default');
  await t.raw.selectOption('#fontFamily', 'mincho');
  await t.sleep(700);

  // 一時保存 → 呼出
  await t.eval(() => window.scrollTo(0, 0));
  await t.click('一時保存');
  await t.sleep(800);
  await t.raw.fill('#saveName', '一学期号_0717');
  await t.sleep(300);
  await t.shot('20-admin-save-modal');
  await t.click('保存する');
  await t.sleep(1200);

  await t.click('呼出');
  await t.sleep(1200);
  await t.shot('21-admin-load-modal');
  await t.click('キャンセル');
  await t.sleep(700);

  // テンプレート
  await t.click('テンプレート');
  await t.sleep(1200);
  await t.raw.fill('#templateName', '学級新聞 標準');
  await t.sleep(300);
  await t.shot('22-admin-template-modal');
  await t.click('保存');
  await t.sleep(1200);
  await t.click('閉じる');
  await t.sleep(700);

  // タグ設定
  await t.click('タグ設定');
  await t.sleep(1200);
  await t.click('タグを追加');
  await t.sleep(500);
  await t.eval(() => {
    const rows = document.querySelectorAll('#tagListContainer .tag-row');
    const r = rows[rows.length - 1];
    const [icon, name, ruby] = r.querySelectorAll('input');
    icon.value = '🎨'; icon.dispatchEvent(new Event('change', { bubbles: true }));
    name.value = '図工'; name.dispatchEvent(new Event('change', { bubbles: true }));
    ruby.value = 'ずこう'; ruby.dispatchEvent(new Event('change', { bubbles: true }));
  });
  await t.sleep(500);
  await t.shot('23-admin-tags-modal');
  await t.click('設定を保存');
  await t.sleep(1200);
  await t.shot('24-admin-toast');

  // ============================================================ 児童の画面に戻る
  await c.raw.reload({ waitUntil: 'networkidle' });
  await c.sleep(2200);
  await c.resize(390, 1800);
  await c.scrollTo('タグ');
  await c.sleep(500);
  await c.shot('25-child-newtag');
  log('児童画面のタグ:', await c.eval(() => [...document.querySelectorAll('.tag-label')].map((e) => e.innerText.replace(/\s+/g, ''))));
};
