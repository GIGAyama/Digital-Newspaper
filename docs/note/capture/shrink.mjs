/* 150KB を超えた画面写真だけを縮める。
 * 縮めるのは「もとの画像から1回だけ」。書き戻したものを次の周でまた縮めると、
 * 倍率が掛け算になって、気づいたら 0.2 倍のぼやけた絵になる（実際にやった）。 */
import { chromium } from 'playwright';
import fs from 'node:fs';
import path from 'node:path';
const DIR = path.resolve(process.argv[2] || 'docs/note/images');
const MAX = 150 * 1024;   // 品質ゲートの ASSET_TOO_BIG と同じ値
const browser = await chromium.launch();
const page = await browser.newPage();
for (const name of fs.readdirSync(DIR).filter((f) => f.endsWith('.png'))) {
  const file = path.join(DIR, name);
  const original = fs.readFileSync(file);
  if (original.length <= MAX) continue;
  const src = 'data:image/png;base64,' + original.toString('base64');
  let best = null, used = 1;
  for (let scale = 0.95; scale >= 0.45; scale -= 0.05) {
    const out = await page.evaluate(async ([s, sc]) => {
      const img = new Image();
      await new Promise((ok, ng) => { img.onload = ok; img.onerror = ng; img.src = s; });
      const c = document.createElement('canvas');
      c.width = Math.round(img.naturalWidth * sc); c.height = Math.round(img.naturalHeight * sc);
      const x = c.getContext('2d');
      x.imageSmoothingQuality = 'high';
      x.drawImage(img, 0, 0, c.width, c.height);
      return c.toDataURL('image/png');
    }, [src, scale]);
    const buf = Buffer.from(out.split(',')[1], 'base64');
    if (buf.length <= MAX) { best = buf; used = scale; break; }
  }
  if (!best) { console.log(name, '縮めても収まらない'); continue; }
  fs.writeFileSync(file, best);
  console.log(name, Math.round(best.length / 1024) + 'KB', '（' + used.toFixed(2) + '倍）');
}
await browser.close();
