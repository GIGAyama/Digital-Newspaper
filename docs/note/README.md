# note 入稿メモ — デジタル・クラス新聞社

連載「教室で使えるかもしれないもの作り」に載せる紹介記事一式です。

```
docs/note/
├── digital-newspaper-note-article.md   記事本文。そのまま note に貼れる形
├── images/                             実際に画面を操作して撮った 17 点
├── capture/                            撮り直すための道具（下記）
└── README.md                           このファイル
```

**2026-08-23 に、GAS 版（v1）の記事から書き直しました。** 旧版の記事と画像
（25点）は、この作りかえの直前のコミット `4e0e528` に残っています。
撮り直す道具も入れ替えました。旧版の `capture/gas-server.mjs` は
`Code.gs` の代役サーバーで、いまのアプリには要りません。

---

## 貼る前に直すところ

**連載番号が `#◯` のままです。** 1行目のタイトルの `#◯` を実際の号数に置きかえてください。
過去の記事が手もとに無く、推測で埋めると連載が壊れるので空けてあります。

**題名は旧版と同じにしてあります。** ポータル（giga-school.com）の各紹介ページは
おたがいの題名を「ほかの紹介」として持ち合っているため、題名を変えると
31 本ぶんを組み直すことになります。中身は全部書き直しましたが、題名だけは
そのまま置いてあります。変えるなら、`tools/build-articles.mjs` を通しで
回せるときにしてください。

---

## note に貼る手順

1. note の新規投稿を開き、記事の1行目（タイトル欄）に `#` を外した見出しを入れる
2. 本文を `## 🏫 はじめに` 以降から貼る。note は `##` を見出しとして受け取り、目次を自動で作る
3. `![...](images/xx.png)` の行は貼っても画像になりません。**その位置で画像を手動アップロードし、Markdown の行は消す**
4. 画像の直後の一文は、note の**キャプション欄**に移す。本文に残すと二重になります
5. 太字・表・記号は使っていないので、貼り付け後の手直しは要りません
6. ハッシュタグの候補は `#GIGAスクール` `#学級新聞` `#WebRTC` `#小学校` `#校務効率化` `#ICT教育`

見出しは 8 本です。目次はこの並びで出ます。

```
🏫 はじめに
📱 このアプリでできること
🔗 あいことばだけで、記事が集まる
🗞️ 縦書きの紙面が、その場で組み上がる
✨ 導入のメリット
🛠️ 【管理者向け】導入手順
📖 【利用者向け】使い方のガイド
📝 まとめ
```

`🔗` と `🗞️` の2本は、文体ガイドの「目玉が既存の見出しに収まらないときだけ
📱 と ✨ のあいだに2本まで足す」に沿って足したものです。

---

## 画像の対応表

| ファイル | 何の画面か | 本文のどこ |
|---|---|---|
| 01-write-empty.png | まっさらな「記事を かく」。記入欄5つ | 📱 |
| 02-write-filled.png | 記入し、タグを選び、写真を入れた状態 | 📱 |
| 03-mine-list.png | わたしの記事の一覧。状態のふだが見える | 📱 |
| 04-room-open.png | へやを開いて、あいことばが出たところ | 🔗 |
| 05-join.png | あいことばを入れるところ | 🔗 |
| 06-sent.png | 送ったあと、ふだが「おくった」に変わった | 🔗 |
| 07-received.png | 届いた記事が一覧に並んだところ | 🔗 |
| 08-published.png | 「のせてもらった」に変わったところ | 🔗 |
| 09-desk-panel.png | 新聞をつくる画面の操作パネル全体 | 🗞️ |
| 10-paper.png | 縦書き3段に組み上がった紙面 | 🗞️ |
| 11-paper-photos.png | 写真を右に回りこませた紙面 | 🗞️ |
| 12-paper-qr.png | QR コードの入った紙面 | 🗞️ |
| 13-paper-blackboard.png | 黒板テーマ＋キウイ丸の紙面 | 🗞️ |
| 14-paper-spring-horizontal.png | 春テーマで横書き2段の紙面 | 🗞️ |
| 15-free-modal.png | 自由記述らんの追加 | 🗞️ |
| 16-paper-free.png | 編集後記の入った紙面 | 🗞️ |
| 17-mobile-desk.png | 390px 幅で新聞を組んでいるところ | 📖 |

いちばん効く1枚は `16-paper-free.png` です。note のヘッダー画像に使うならこれです。

**画像はすべて 150KB 以内に収めてあります。** ポータル（giga-school.com）の紹介ページは
これらをこのリポジトリのドメインから直に読むので、重いと表示が遅くなります。
品質ゲート（`npm run check`）の `ASSET_TOO_BIG` が見張っています。

---

## 撮影について、書き手が知っておくべきこと

**画面はすべて本物です。`index.html` をそのままブラウザで開いて、
文字を打ち、ボタンを押して撮りました。合成も加工もしていません。**
2台の端末で通しています。片方が記事を4本書いてへやを開き、もう片方が
2本書いて入り、送り、それを含めて紙面を組む、という順です。

そのうえで、次の3つは撮影のために用意したものです。記事の中でも
「実際の教室での出来事」としては一切書いていません。

1. **記事の中身は見本です。** 三年二組の6本の記事と児童名は、撮影のために書いたもので、
   実在の学級のものではありません
2. **写真は絵です。** 児童の顔が写った本物の写真は使えないので、
   運動会・ミニトマト・プールの3点を、その場で `<canvas>` に描いたものを
   選ばせています（`capture/shots.mjs` の `DRAW`）
3. **差し替えたのは2箇所だけです。** アプリの HTML と JavaScript には手を入れていません
   - Google Fonts の読み込み先を、npm から取った同じ書体のローカル写しに向けた
     （作業環境から `fonts.googleapis.com` に出られないため）
   - `window.Peer` の既定の接続先を、手もとに立てた PeerServer に向けた
     （`0.peerjs.com` に出られないため）

**QR コードは本物です。** 記事に写っている QR は、実際に qr-creator が描いたものです。

---

## 記事に書いた数字の出どころ

| 記事の記述 | 出どころ |
|---|---|
| 記入欄が5つ | `index.html` のフォーム。タイトル・本文・記者名・タグ・写真 |
| ふだは3種類 | `statusBadge()`。したがき／おくった／のせてもらった |
| あいことばは十文字 | `CODE_LENGTH = 10` |
| 0 O 1 I L を使わない | `CODE_ALPHABET = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'` |
| 八十二兆通り | 31 の 10 乗。`CODE_ALPHABET.length ** CODE_LENGTH` |
| 小文字でも区切りなしでも通る | `normalizeRoomCode()` |
| 届いた合図が返ってはじめて「おくった」 | `onGuestData()` の `ack` の分岐。`sendArticle()` では立てない |
| 届いた記事はチェックが入っていない | `onHostData()` が `upsertArticle(a, false)` を呼ぶ |
| 自分の記事ははじめからチェックが入る | `upsertArticle()` の既定は `isMine(a)` |
| 長辺八百ピクセル | `PHOTO_MAX_EDGE = 800` |
| 六十から百キロバイト | AUDIT.md 第4回。実測 |
| 写真つきで四十本くらい | `localStorage` 5MB ÷ 100KB。README.md の「写真は端末の中で小さくしてから持ちます」 |
| 八割で赤くなる | `updateStorageMeter()` の `pct >= 80` |
| QR は見た目の倍の画素 | `qrScale()` は `Math.max(2, ...)` |
| 読み取れない白い四角が刷られた | AUDIT.md 第2回。QRious を cdnjs から読んでいた版の実測（塗られた画素 2500 → 0） |
| テーマ7種類 | `#themeSelect` の選択肢 |
| フォント5種類 | `#fontFamily` の選択肢 |
| 見出しの飾り6種類 | `#titleStyle` の選択肢 |
| 段は1段から4段 | `#colCount` の選択肢 4 つ |
| 文字の大きさは8から24 | `#fontSizeRange` の `min=8 max=24` |
| 紙面はA4一枚ぶん | `.page-a4-v` は 210mm×297mm、`#paperArea` は `overflow: hidden` |
| 通信がなくても起動して印刷までできる | AUDIT.md 第4回。Service Worker 登録後に回線を切って実測 |
| 390px で操作パネルが縦に積まれる | AUDIT.md 第4回。横はみ出し 0 を実測 |

記事に書いた数字はすべてこの表のどれかに当たります。書き足すときは、
コードか AUDIT.md のどちらかで裏を取ってから足してください。

---

## 撮り直すときの手順

`capture/` に道具一式が入っています。**GAS 版と違い、代役のサーバーは要りません。**
アプリそのものを配って撮ります。

```bash
npm i --no-save playwright peer \
  @fontsource/shippori-mincho @fontsource/noto-sans-jp \
  @fontsource/yomogi @fontsource/kiwi-maru @fontsource/biz-udpgothic
npx playwright install chromium

node docs/note/capture/prepare.mjs .capture     # 撮影用の写しを作る
node docs/note/capture/peerserver.cjs &          # 手もとのシグナリングサーバー
node docs/note/capture/serve.cjs .capture 8098 & # 撮影用の写しを配る
LANG=ja_JP.UTF-8 LANGUAGE=ja node docs/note/capture/shots.mjs docs/note/images
node docs/note/capture/shrink.mjs docs/note/images   # 150KB を超えた分だけ縮める
```

`LANGUAGE=ja` を落とさないでください。落とすとブラウザの部品が英語になり、
日付の絞りこみが `mm/dd/yyyy` になります。日本語の記事の画像としては具合が悪いところです。

フォントを npm から入れているのは、この作業環境から Google Fonts に出られないためです。
出られる環境なら `prepare.mjs` に `--keep-fonts` を渡してください（差し替えません）。

画面ごとの寸法は `shots.mjs` に書いてあります。紙面は要素だけを撮っているので、
A4 縦なら 794×1123px です。150KB を超えたものだけ `shrink.mjs` が縮めます。
**縮めるのはもとの画像から1回だけです。** 書き戻したものをまた縮めると
倍率が掛け算になって、気づいたら 0.2 倍のぼやけた絵になります（実際にやりました）。
