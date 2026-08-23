/**
 * 📰 デジタル・クラス新聞社 システムコード (Ver.1.0 Release)
 * * 概要:
 * 児童が記事を投稿し、教師がそれを編集して学級新聞を作成するWebアプリケーションです。
 * Google Sheetsをデータベース、Google Driveを画像ストレージとして使用します。
 * * 主な機能:
 * - 記事投稿 (画像添付、タグ付け)
 * - 記事管理 (一覧表示、フィルタリング、編集)
 * - 新聞レイアウト作成 (縦書き/横書き、段組み、自動組版)
 * - 設定保存 (レイアウト状態、デザインテンプレート)
 * - タグ管理 (カスタマイズ可能なタグ設定)
 */

// ==================================================
// 1. 設定・定数 (Configuration)
// ==================================================

/** プロパティストアのキー */
const KEYS = {
  FOLDER_ID: 'IMAGE_FOLDER_ID',      // 画像保存先フォルダID
  TAG_SETTINGS: 'TAG_SETTINGS',      // タグ設定JSON
  TEACHER_EMAILS: 'TEACHER_EMAILS',  // 管理画面を使える先生のメール（カンマ区切り）
  OWNER_EMAIL: 'OWNER_EMAIL'         // TEACHER_EMAILS 未設定時の照合先（初回セットアップ実行者）
};

/** シート名定義 */
const SHEETS = {
  ARTICLES: 'Articles',      // 記事データ
  SYSTEM: 'SystemData'       // 設定・テンプレートデータ
};

/** デフォルトのタグ設定 (初期化用) */
const DEFAULT_TAGS = [
  { icon: "📰", name: "ニュース", ruby: "ニュース" },
  { icon: "🎌", name: "行事", ruby: "ぎょうじ" },
  { icon: "✏️", name: "学習", ruby: "がくしゅう" },
  { icon: "⚽", name: "遊び", ruby: "あそび" },
  { icon: "🍀", name: "その他", ruby: "そのた" }
];

// ==================================================
// 1.6 シートの作り（スキーマ）— 自動の点検と、安全な修整
// ==================================================
//
// このアプリは「スプレッドシートのコピー」で配る（README「配り方」）。
// コピーしたあとのファイルは先生のものなので、列を足す・並べ替える・
// シートを消す・見出しを打ち直す、が実際に起きる。
// 起きたときに **何も言わずに壊れる** のがいちばん困るので、3段に分けた。
//
//   1. ensureSchema_(ss)  … 開くたびに、足りないシートと見出し行だけを自動でそろえる
//   2. checkSchema_(ss)   … どこがどう違うかを数え上げる。**何も書き換えない**
//   3. repairSchema_(ss)  … メニューから、安全に直せるものだけを直す
//
// ⚠️ 修整でやるのは「足す」と「書き方をそろえる」だけで、消す・動かすは一切しない。
//    見出しがずれているとき、見出しだけを正しく書き換えると
//    **間違った列に正しいラベルが付き、事故が見えなくなる。**
//    列の中身を動かす判断は人がする。ここは「どこがどうずれているか」を言うだけにする。
//
// なぜ「足す」だけで足りるか: 読み書きはすべて見出し名で行う（列番号を使わない）。
// だから列を並べ替えられても動く。足りないのは「その名前の列が無い」ときだけで、
// 右端に足せば、既存のデータを 1 セルも動かさずに直る。

/**
 * このアプリが使うシートと、その見出し行。ここが唯一の正。
 * 列を足すときは、ここに足すだけでよい（配ったコピーにも次に開いたときそろう）。
 */
const SCHEMA_ = [
  {
    name: SHEETS.ARTICLES,
    header: ['ID', 'Title', 'Body', 'ImageURL', 'Reporter', 'Timestamp', 'Status', 'Tag'],
    hidden: false
  },
  {
    name: SHEETS.SYSTEM,
    header: ['Type', 'Name', 'Data', 'Date'],
    hidden: true
  }
];

/** 見出しの照合用に、前後の空白・全角空白・大文字小文字・アンダースコアの差を落とす */
function schemaKey_(name) {
  return String(name == null ? '' : name).replace(/[\s　_]/g, '').toLowerCase();
}

/** 1 行目を文字列の配列で読む。空のシートでは空配列。 */
function readHeaderRow_(sheet) {
  const width = sheet.getLastColumn();
  if (width < 1) return [];
  return sheet.getRange(1, 1, 1, width).getValues()[0]
    .map(v => (v === null || v === undefined) ? '' : String(v).trim());
}

/** 空のシートに見出しを書く。データのあるシートには絶対に使わない。 */
function writeHeader_(sheet, spec) {
  sheet.getRange(1, 1, 1, spec.header.length).setValues([spec.header]);
  sheet.setFrozenRows(1);
}

/**
 * 足りないシートと、空のシートの見出しだけを作る。
 *
 * ふつうは 1 枚も足りないので、その場合はロックを取らずに帰る
 * （40 台が一斉に開く時間に、全員がロック待ちに並ぶのを避けるため）。
 * ロックが取れなかったときも、あるものだけで進める。書き込む側が
 * articleColumns_() で改めて確かめるので、ここで止める必要はない。
 */
function ensureSchema_(ss) {
  const todo = SCHEMA_.filter(spec => {
    const sheet = ss.getSheetByName(spec.name);
    return !sheet || sheet.getLastRow() === 0;
  });
  if (todo.length === 0) return ss;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    console.warn('シートをそろえるロックが取れませんでした（あるものだけで進めます）: ' + e);
    return ss;
  }
  try {
    SCHEMA_.forEach(spec => {
      let sheet = ss.getSheetByName(spec.name);
      if (sheet && sheet.getLastRow() > 0) return;   // ロック待ちの間に誰かが作っていた
      if (!sheet) sheet = ss.insertSheet(spec.name);
      writeHeader_(sheet, spec);
      // 非表示にするのは作った直後の 1 回だけ。先生が自分で表示に戻したものを
      // 開くたびに隠し返すと、先生の操作を毎回取り消すことになる。
      if (spec.hidden) {
        try { sheet.hideSheet(); } catch (e2) { /* 表示中のシートが1枚だけのときは隠せない */ }
      }
    });
  } finally {
    lock.releaseLock();
  }
  return ss;
}

/**
 * シートの作りが SCHEMA_ のとおりかを数え上げる。**何も書き換えない。**
 *
 * @return {{sheet:string, kind:string, detail:string, fix:string, col:number, to:string}[]}
 *   fix は repairSchema_ が安全に直せる種類のときだけ入る。
 *   'sheet'（シートを作る）/ 'header'（空のシートに見出しを書く）/
 *   'rename'（見出しの書き方をそろえる）/ 'column'（右端に列を足す）。
 *   空文字のものは「人が確かめること」で、自動では触らない。
 */
function checkSchema_(ss) {
  const found = [];

  SCHEMA_.forEach(spec => {
    const sheet = ss.getSheetByName(spec.name);
    if (!sheet) {
      found.push({
        sheet: spec.name, kind: 'シートが無い',
        detail: '「' + spec.name + '」シートがありません', fix: 'sheet', col: -1, to: ''
      });
      return;
    }

    const header = readHeaderRow_(sheet);
    if (header.length === 0 || header.join('') === '') {
      found.push({
        sheet: spec.name, kind: '見出しが無い', detail: '1 行目が空です',
        // 中身が 1 行も無いときだけ書いてよい。下にデータがあるなら、
        // それは「見出しを消した」状態なので人に確かめてもらう。
        fix: sheet.getLastRow() === 0 ? 'header' : '', col: -1, to: ''
      });
      return;
    }

    const at = {};          // 正確に一致した見出し名 -> 何列目（0 起点）
    const dup = [];
    header.forEach((name, i) => {
      if (!name) return;
      if (at[name] === undefined) at[name] = i; else dup.push(name);
    });

    const loose = {};       // 書き方の差を落とした見出し -> 何列目（0 起点）
    header.forEach((name, i) => {
      const k = schemaKey_(name);
      if (k && loose[k] === undefined) loose[k] = i;
    });

    // 1 行目がそもそも見出しに見えないときは、1 か所も直さない。
    // 見出しを消して詰めた（＝1 行目がデータになった）状態でここを直すと、
    // データの上に見出しを上書きして 1 件消すことになる。
    const hits = spec.header.filter(h => loose[schemaKey_(h)] !== undefined).length;
    if (hits === 0) {
      found.push({
        sheet: spec.name, kind: '1 行目が見出しに見えない',
        detail: '1 行目は「' + header.slice(0, 3).join('／') + '」でした。見出しの行ごと消えた可能性があります。'
          + '1 行目に行を挿入して「' + spec.header.join('／') + '」を戻してください'
          + '（データを動かすことになるので、自動では直しません）',
        fix: '', col: -1, to: ''
      });
      return;
    }

    spec.header.forEach(name => {
      if (at[name] !== undefined) return;
      const near = loose[schemaKey_(name)];
      if (near !== undefined) {
        found.push({
          sheet: spec.name, kind: '見出しの書き方がちがう',
          detail: (near + 1) + ' 列目が「' + header[near] + '」になっています（正しくは「' + name + '」）',
          fix: 'rename', col: near, to: name
        });
      } else {
        found.push({
          sheet: spec.name, kind: '列が足りない',
          detail: '「' + name + '」の列がありません',
          fix: 'column', col: -1, to: name
        });
      }
    });

    // 並べ替えは咎めない。読み書きは見出し名で行うので、このままで動く。
    // ただし黙っているとこちらの想定と違うことに気づけないので、出しておく。
    const order = spec.header.filter(n => at[n] !== undefined).map(n => at[n]);
    const sorted = order.slice().sort((a, b) => a - b);
    if (order.join(',') !== sorted.join(',')) {
      found.push({
        sheet: spec.name, kind: '列の並びが既定と違う',
        detail: '読み書きは見出し名で行うので、このままでも動きます（直しません）',
        fix: '', col: -1, to: ''
      });
    }

    const known = {};
    spec.header.forEach(n => { known[schemaKey_(n)] = true; });
    const extra = header.filter(n => n && !known[schemaKey_(n)]);
    if (extra.length > 0) {
      found.push({
        sheet: spec.name, kind: 'アプリが知らない列がある',
        detail: '「' + extra.join('」「') + '」。消しません。記事が届いても、この列は空のままです',
        fix: '', col: -1, to: ''
      });
    }

    if (dup.length > 0) {
      found.push({
        sheet: spec.name, kind: '同じ見出しが 2 つある',
        detail: '「' + dup.join('」「') + '」。どちらを読むか決められないので、自動では直しません',
        fix: '', col: -1, to: ''
      });
    }
  });

  return found;
}

/** 右端に列を 1 本足して見出しを書く。既存のセルは 1 つも動かさない。@return 足した列の番号（1 起点） */
function appendColumn_(sheet, name) {
  const col = sheet.getLastColumn() + 1;
  const max = sheet.getMaxColumns();
  if (col > max) sheet.insertColumnsAfter(max, col - max);
  sheet.getRange(1, col).setValue(name);
  return col;
}

/**
 * checkSchema_ が「安全に直せる」と判断したものだけを直す。
 * 消す・動かす・並べ替えるは一切しない。
 *
 * @return {string[]} 実際にやったことの一覧（何もしなかったときは空）
 */
function repairSchema_(ss) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('ほかの処理と重なっています。少し待ってから、もう一度お試しください。');
  }
  try {
    const done = [];
    // ロックを取ってから数え直す。待っている間に誰かが直しているかもしれない。
    checkSchema_(ss).forEach(f => {
      if (!f.fix) return;
      const spec = SCHEMA_.filter(s => s.name === f.sheet)[0];
      if (!spec) return;

      if (f.fix === 'sheet') {
        const sheet = ss.insertSheet(spec.name);
        writeHeader_(sheet, spec);
        if (spec.hidden) {
          try { sheet.hideSheet(); } catch (e2) { /* 表示中のシートが1枚だけのときは隠せない */ }
        }
        done.push('「' + spec.name + '」シートを作り、見出しを書きました');
        return;
      }

      const sheet = ss.getSheetByName(f.sheet);
      if (!sheet) return;

      if (f.fix === 'header') {
        writeHeader_(sheet, spec);
        done.push('「' + f.sheet + '」の 1 行目に見出しを書きました');
      } else if (f.fix === 'rename') {
        sheet.getRange(1, f.col + 1).setValue(f.to);
        done.push('「' + f.sheet + '」' + (f.col + 1) + ' 列目の見出しを「' + f.to + '」にそろえました');
      } else if (f.fix === 'column') {
        const col = appendColumn_(sheet, f.to);
        done.push('「' + f.sheet + '」の ' + col + ' 列目に「' + f.to + '」の列を足しました（中身は空です）');
      }
    });
    return done;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 記事シートの見出し行と、見出し名 -> 列番号（1 起点）の対応を返す。
 *
 * 足りない見出しがあれば右端に足す。書き込みは必ずこの対応表を通すので、
 * 先生が列を並べ替えていても正しい列に入る。
 *
 * ⚠️ **呼ぶ側はロックを持っていないこと。** 列を足すときだけ、ここで自分で
 *    短いロックを取り、取れてから見出しをもう一度読む。持ったまま呼ぶと
 *    二重取得になり、持たずに足すと 2 人ぶんが同時に足して「Tag」列が 2 本できる。
 */
function articleColumns_(sheet) {
  const spec = SCHEMA_.filter(s => s.name === SHEETS.ARTICLES)[0];
  const missing = (header) =>
    spec.header.filter(name => !header.some(n => schemaKey_(n) === schemaKey_(name)));

  let header = readHeaderRow_(sheet);

  // 1 行目が見出しに見えないときは足さない。データの右に空列を積むだけになる。
  const hits = spec.header.length - missing(header).length;
  if (hits > 0 && missing(header).length > 0) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
    } catch (e) {
      // 足せなくても、あるものだけで書く。足りない列の中身は落ちるが、
      // 記事そのものは残る。点検メニューで気づける。
      console.warn('足りない列を足すロックが取れませんでした: ' + e);
      return columnMap_(header);
    }
    try {
      header = readHeaderRow_(sheet);          // 待っている間に誰かが足したかもしれない
      missing(header).forEach(name => {
        appendColumn_(sheet, name);
        header = readHeaderRow_(sheet);
      });
    } finally {
      lock.releaseLock();
    }
  }

  return columnMap_(header);
}

/** 見出しの並びから、見出し名 -> 列番号（1 起点）の対応を作る */
function columnMap_(header) {
  const map = {};
  header.forEach((name, i) => {
    const k = schemaKey_(name);
    if (k && map[k] === undefined) map[k] = i + 1;
  });
  return { header: header, col: (name) => map[schemaKey_(name)] || 0 };
}

// ==================================================
// 1.5 教員判定 (Authorization)
// ==================================================
//
// 管理画面は URL を知っていれば誰でも開けてしまう作りだった。
// 画面を隠すだけでは足りない（google.script.run から関数を直接呼べる）ので、
// 「画面の入口」と「管理系のサーバー関数」の両方で同じ判定を通す。
//
// ── ウェブアプリは「自分（デプロイした先生）」として動く ─────────────
//
// appsscript.json は `executeAs: USER_DEPLOYING` / `access: DOMAIN`。
// つまりサーバー側の処理は、**誰が開いても、デプロイした先生の権限で走る。**
// おかげで児童は先生のスプレッドシートにもドライブにも一切アクセス権が要らない
// （＝児童がシートを開いて onOpen を動かすことも、記録を直接いじることもできない）。
//
// ⚠️ そのぶん、**身元の判定に Session.getEffectiveUser() を使ってはいけない。**
//    この形では effective user は児童が開いても「デプロイした先生」になるので、
//    使った瞬間に学級全員が先生として通る。
//    身元は必ず Session.getActiveUser()（＝いま画面を見ている人）だけで決める。
//    同一 Workspace ドメイン内なら、この形でも本人のメールが取れる
//    （access: DOMAIN で校内に限っているのはそのためでもある）。
//    取れなければ空文字になり、下の判定は false に倒れる（fail-closed）。

/** 先生以外に見せる文言。メニューの入口で共通に使う。 */
const TEACHER_ONLY_MSG_ =
  '⚠️ このアカウントでは、この操作はできません。\n' +
  'メニューの「2. 先生のメールアドレスを設定」で、いま使っているアカウントを登録してください。';

/** 文字列を照合用に正規化する（前後の空白と大文字小文字の差を無くす） */
function normalizeEmail_(value) {
  return String(value == null ? '' : value).trim().toLowerCase();
}

/**
 * 管理画面を使える人のメール一覧を返す。
 * TEACHER_EMAILS（カンマ区切り）が正。未設定のときだけ、
 * 初回セットアップ実行者として記録した OWNER_EMAIL を使う。
 */
function getTeacherEmails_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(KEYS.TEACHER_EMAILS) || '';
  const list = raw.split(',').map(normalizeEmail_).filter(v => v);
  if (list.length > 0) return list;

  const owner = normalizeEmail_(props.getProperty(KEYS.OWNER_EMAIL));
  return owner ? [owner] : [];
}

/**
 * いま画面を見ている人が先生かどうか。
 * メールが取れないとき（ドメイン外・未ログイン・共有アカウント）は必ず false。
 * 許可リストが空のときも false（設定するまでは誰も入れない＝安全側に倒す）。
 *
 * ⚠️ ここを getEffectiveUser() に替えないこと。ウェブアプリは
 *    「デプロイした先生」として走るので、児童が開いても true になる。
 */
function isTeacher_() {
  let email = '';
  try {
    email = normalizeEmail_(Session.getActiveUser().getEmail());
  } catch (e) {
    console.warn('利用者のメールを取得できませんでした: ' + e);
    return false;
  }
  if (!email) return false;

  const allowed = getTeacherEmails_();
  if (allowed.length === 0) {
    console.warn('TEACHER_EMAILS も OWNER_EMAIL も未設定のため、管理機能を拒否しました。');
    return false;
  }
  return allowed.indexOf(email) !== -1;
}

/** 管理系のサーバー関数の先頭で呼ぶ。先生でなければ例外を投げて処理を止める。 */
function requireTeacher_() {
  if (!isTeacher_()) {
    throw new Error('この操作は先生のアカウントでしか実行できません。学校のアカウントでログインし直してください。');
  }
}

/**
 * 初回セットアップ（スプレッドシートのメニュー操作）の実行者を OWNER_EMAIL として記録する。
 * TEACHER_EMAILS が設定されていれば何もしない。
 *
 * ⚠️ ここだけは getEffectiveUser() を使う。**メニューからしか呼ばない**からである。
 *    メニューはコンテナバインドの文脈で動き、そこでの実行者は
 *    「そのスプレッドシートを開いている人」＝ファイルの持ち主の先生になる。
 *    ウェブアプリの文脈（＝児童が開いた画面）からは、呼び出し元がすべて
 *    SpreadsheetApp.getUi() で止まるうえ、仮に通っても effective user は
 *    デプロイした先生自身なので、他人が先生になることはない。
 *
 *    **この関数をウェブアプリから呼べる場所へ移さないこと。**
 */
function ensureOwnerEmail_() {
  const props = PropertiesService.getScriptProperties();
  if ((props.getProperty(KEYS.TEACHER_EMAILS) || '').trim()) return '';
  const existing = normalizeEmail_(props.getProperty(KEYS.OWNER_EMAIL));
  if (existing) return existing;

  let email = '';
  try {
    email = normalizeEmail_(Session.getEffectiveUser().getEmail());
  } catch (e) {
    email = '';
  }
  if (!email) {
    console.warn('セットアップ実行者のメールが取得できず、OWNER_EMAIL を記録できませんでした。');
    return '';
  }
  props.setProperty(KEYS.OWNER_EMAIL, email);
  console.info('OWNER_EMAIL を記録しました（管理画面を使えるアカウント）。');
  return email;
}

/** メニューから、管理画面を使える先生のメールを登録する */
function setTeacherEmails() {
  const ui = SpreadsheetApp.getUi();      // 画面が無ければ、ここで止まる
  const props = PropertiesService.getScriptProperties();
  ensureOwnerEmail_();
  // getUi() が例外になることを防御と数えない（v5 §5-1）。2 枚目の判定を必ず通す。
  // 初回は直前の ensureOwnerEmail_ が実行者を記録するので、ここは通る。
  if (!isTeacher_()) { ui.alert(TEACHER_ONLY_MSG_); return; }
  const current = props.getProperty(KEYS.TEACHER_EMAILS) || getTeacherEmails_().join(',');

  const result = ui.prompt(
    '管理画面を使える先生の設定',
    '先生のメールアドレスをカンマ区切りで入力してください：\n(現在: ' + (current || '未設定') + ')',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const list = result.getResponseText().split(',').map(normalizeEmail_).filter(v => v);
  if (list.length === 0) {
    ui.alert('❌ 1件も入力されていません。変更をキャンセルしました。');
    return;
  }
  props.setProperty(KEYS.TEACHER_EMAILS, list.join(','));
  ui.alert('✅ ' + list.length + ' 件のアカウントを登録しました。\nこのアカウント以外は編集室に入れません。');
}

/**
 * このスクリプトが束ねられているスプレッドシート。
 *
 * このアプリは **コンテナバインド** で配る。スプレッドシートのコピーを配り、
 * そのファイルにこのスクリプトが束ねられているので、シート ID を持つ必要も、
 * 自動生成する必要も無い。先生が開いているそのファイルが、そのまま中身である。
 *
 * ⚠️ 独立スクリプト（script.new で作ったもの）に貼り付けると、ここが null になる。
 *    その場合に openById や create で自分を救おうとしてはいけない
 *    （児童一人ひとりの権限で走るので、権限の無い子が 1 回開くだけで
 *    学級のデータが入ったシートから空のシートへ差し替わる。画面には何も出ない）。
 *    直し方を書いて止める。
 */
function getDb_() {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    ss = null;
  }
  if (!ss) {
    throw new Error(
      'このスクリプトがスプレッドシートに束ねられていません。' +
      '案内ページのとおりスプレッドシートをコピーして、そのコピーの「拡張機能 → Apps Script」から公開し直してください。'
    );
  }
  return ensureSchema_(ss);
}

// ==================================================
// 2. スプレッドシート連携・メニュー (Spreadsheet UI)
// ==================================================

/**
 * スプレッドシートを開いたときのメニュー。
 * コンテナバインドのときだけ意味がある（ウェブアプリとして動くときは呼ばれない）。
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📰 新聞システム')
      .addItem('1. 写真フォルダの設定', 'setFolderId')
      .addItem('2. 先生のメールアドレスを設定', 'setTeacherEmails')
      .addSeparator()
      .addItem('3. 先生用管理画面を開く', 'showAdminUrl')
      .addSeparator()
      .addItem('4. シートを点検する', 'showSheetCheck')
      .addItem('5. シートを直す（足りないものを足す）', 'repairSheets')
      .addToUi();
  } catch (e) {
    // 画面が無い文脈では何もしない
  }
}

/** 点検結果を箇条書きの文字列にする */
function formatSchemaFindings_(found) {
  return found.map(f => '・「' + f.sheet + '」' + f.kind + '：' + f.detail).join('\n');
}

/**
 * メニュー「4. シートを点検する」。見つけたものを並べるだけで、**何も書き換えない。**
 *
 * ⚠️ google.script.run は末尾 `_` の無い関数を誰でも直接呼べる。児童からも呼ばれる前提で
 *    2 枚重ねてある。(1) 先に getUi() を取る（ウェブアプリの文脈には画面が無いので、
 *    シートを 1 枚も読まずにここで例外になる）。(2) そのうえで先生かどうかを見る。
 *    戻り値は無く、返すのは見出しの並びだけなので、児童の記事や名前は出ない。
 */
function showSheetCheck() {
  const ui = SpreadsheetApp.getUi();      // 画面が無ければ、ここで止まる
  ensureOwnerEmail_();
  if (!isTeacher_()) { ui.alert(TEACHER_ONLY_MSG_); return; }

  const found = checkSchema_(getDb_());
  if (found.length === 0) {
    ui.alert('シートの点検', 'シートの作りは想定どおりです。', ui.ButtonSet.OK);
    return;
  }
  const fixable = found.filter(f => f.fix);
  ui.alert(
    'シートの点検',
    '次のところが、アプリの想定と違います。\n\n' + formatSchemaFindings_(found) +
    (fixable.length > 0
      ? '\n\nこのうち ' + fixable.length + ' 件は「5. シートを直す」で直せます（足すだけで、消したり動かしたりはしません）。'
      : '\n\n自動で直せるものはありません。上の内容を確かめて、手で直してください。'),
    ui.ButtonSet.OK
  );
}

/**
 * メニュー「5. シートを直す」。安全に直せるものだけを、確認を取ってから直す。
 * 直すのは「シートを作る」「空のシートに見出しを書く」「見出しの書き方をそろえる」
 * 「右端に足りない列を足す」の 4 つだけ。**消す・動かすは一切しない。**
 *
 * 認可の作りは showSheetCheck と同じ（getUi() → 先生かどうか）。
 */
function repairSheets() {
  const ui = SpreadsheetApp.getUi();      // 画面が無ければ、ここで止まる
  ensureOwnerEmail_();
  if (!isTeacher_()) { ui.alert(TEACHER_ONLY_MSG_); return; }

  const ss = getDb_();
  const found = checkSchema_(ss);
  const fixable = found.filter(f => f.fix);
  const manual = found.filter(f => !f.fix);

  if (fixable.length === 0) {
    ui.alert('シートの修整',
      found.length === 0
        ? 'シートの作りは想定どおりです。直すところはありません。'
        : '自動で直せるものはありませんでした。次のところは手で確かめてください。\n\n' + formatSchemaFindings_(manual),
      ui.ButtonSet.OK);
    return;
  }

  const answer = ui.alert('シートの修整',
    '次のとおり直します。列を消したり動かしたりはしません。\n\n' + formatSchemaFindings_(fixable) +
    '\n\n実行してよろしいですか。',
    ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;

  const done = repairSchema_(ss);
  ui.alert('シートの修整',
    (done.length > 0 ? '次のとおり直しました。\n\n' + done.map(d => '・' + d).join('\n') : '直すところはありませんでした。') +
    (manual.length > 0
      ? '\n\n次のところは自動では直しません。手で確かめてください。\n' + formatSchemaFindings_(manual)
      : ''),
    ui.ButtonSet.OK);
}

function setFolderId() {
  const ui = SpreadsheetApp.getUi();      // 画面が無ければ、ここで止まる
  // 初回セットアップ実行者を管理者として記録する（TEACHER_EMAILS 未設定時のみ）
  ensureOwnerEmail_();
  // getUi() が例外になることを防御と数えない（v5 §5-1）。2 枚目の判定を必ず通す。
  if (!isTeacher_()) { ui.alert(TEACHER_ONLY_MSG_); return; }
  const props = PropertiesService.getScriptProperties();
  const currentId = props.getProperty(KEYS.FOLDER_ID) || '';

  const result = ui.prompt(
    '写真保存用フォルダの設定',
    'GoogleドライブのフォルダIDを入力してください：\n(現在: ' + (currentId ? currentId : '未設定/自動生成') + ')',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const newId = result.getResponseText().trim();
    if (!newId) {
      ui.alert('❌ IDが空です。変更をキャンセルしました。');
      return;
    }
    try {
      DriveApp.getFolderById(newId);
      props.setProperty(KEYS.FOLDER_ID, newId);
      ui.alert('✅ 設定を保存しました！\n以降の投稿画像はこのフォルダに保存されます。');
    } catch (e) {
      ui.alert('⚠️ エラー: 指定されたIDのフォルダが見つかりません。\n権限があるか、IDが間違っていないか確認してください。');
    }
  }
}

function showAdminUrl() {
  const ui = SpreadsheetApp.getUi();
  ensureOwnerEmail_();

  if (!isTeacher_()) { ui.alert(TEACHER_ONLY_MSG_); return; }

  const url = ScriptApp.getService().getUrl();

  if (!url) {
    ui.alert('⚠️ エラー: WebアプリのURLが取得できません。\nまず「デプロイ」→「新しいデプロイ」を実行して、Webアプリとして公開してください。');
    return;
  }

  const htmlOutput = HtmlService.createHtmlOutput(
    '<div style="text-align:center; padding:20px; font-family:sans-serif; color:#333;">' +
    '<h3 style="margin-top:0;">新聞編集室へのアクセス</h3>' +
    '<p>以下のボタンから管理画面へ移動できます。</p>' +
    '<a href="' + url + '?p=admin" target="_blank" style="background:#007bff; color:white; padding:12px 25px; text-decoration:none; border-radius:5px; font-weight:bold; display:inline-block; box-shadow:0 2px 5px rgba(0,0,0,0.2);">🚀 編集室に入る</a>' +
    '<p style="margin-top:20px; font-size:0.85rem; color:#666;">または投稿画面へ：<br><a href="' + url + '" target="_blank" style="color:#007bff;">📝 記者投稿ポスト</a></p>' +
    '</div>'
  ).setWidth(400).setHeight(280);

  ui.showModalDialog(htmlOutput, '管理画面へのアクセス');
}

// ==================================================
// 3. Webアプリ エントリーポイント (DoGet)
// ==================================================

function doGet(e) {
  const page = (e && e.parameter) ? e.parameter.p : '';
  const teacher = isTeacher_();
  let template;
  let title;

  if (page === 'admin') {
    // 画面を隠すだけでは足りないが、入口も塞ぐ（サーバー関数側でも同じ判定を通す）
    if (!teacher) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:sans-serif; padding:24px; color:#333; line-height:1.8;">' +
        '<h2 style="margin-top:0;">🔒 このページは先生専用です</h2>' +
        '<p>編集室は、学校のアカウントで登録された先生だけが開けます。</p>' +
        '<p style="font-size:0.9rem; color:#666;">先生へ：スプレッドシートのメニュー「📰 新聞システム」→「2. 先生のメールアドレスを設定」から、使うアカウントを登録してください。</p>' +
        '</div>'
      ).setTitle('デジタル新聞編集室');
    }
    template = HtmlService.createTemplateFromFile('admin');
    title = 'デジタル新聞編集室';
  } else {
    // 児童用の投稿ページは post.html。
    // ★ 名前を index から post に替えてある。リポジトリ直下の index.html は
    //   GitHub Pages（digital-newspaper.giga-school.com）の導入案内ページで、
    //   GAS へは送らない（.claspignore）。同じ名前のままだと、案内ページを
    //   児童の投稿画面として配ることになる。
    template = HtmlService.createTemplateFromFile('post');
    title = 'デジタルクラス新聞社';
  }
  // 児童が見る画面に編集室への入口を出さないための目印（post.html で使う）
  template.isTeacher = teacher;

  return template.evaluate()
    .setTitle(title)
    // viewport-fit=cover が無いと、切り欠きのある端末で env(safe-area-inset-*) が使えない。
    // GAS は画面を iframe で包むため、HTML 側の <meta> だけでは足りず、
    // サーバー側で足すこのタグにも要る（v5 §5）。前回は HTML 側だけを直して漏れていた。
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=1teNY1fFaXzZC3HivQIQ4t9nu49xYsbJN&.png');
}

// ==================================================
// 4. データ処理ロジック (Data Logic)
// ==================================================

// --- 写真の共有設定 (Photo Sharing) ---

/**
 * 児童の写真は「リンクを知っている全員」には公開しない。
 * まず学校ドメイン内だけに限定し、Google Workspace 以外のアカウント
 * （個人の Gmail など）で DOMAIN_WITH_LINK が使えないときは PRIVATE まで閉じる。
 * どちらになったかは必ずログに残す。
 * @return {string} 実際に設定できた共有範囲
 */
function applyPhotoSharing_(file) {
  try {
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    console.info('写真の共有範囲を DOMAIN_WITH_LINK にしました: ' + file.getId());
    return 'DOMAIN_WITH_LINK';
  } catch (e) {
    console.warn(
      'DOMAIN_WITH_LINK を設定できませんでした（Google Workspace アカウントではない可能性）。' +
      'PRIVATE にフォールバックします: ' + e
    );
  }

  try {
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    console.warn('写真の共有範囲を PRIVATE にしました: ' + file.getId());
  } catch (e2) {
    // ここまで失敗したら共有範囲は既定（作成者のみ）のまま。写真は残す。
    console.error('写真の共有範囲を変更できませんでした: ' + file.getId() + ' / ' + e2);
  }

  // PRIVATE のままだと先生が記事の写真を見られないので、登録済みの先生にだけ閲覧を許す
  shareWithTeachers_(file);
  return 'PRIVATE';
}

/** 登録済みの先生にだけ閲覧権を渡す（PRIVATE フォールバック時の救済） */
function shareWithTeachers_(file) {
  const teachers = getTeacherEmails_();
  if (teachers.length === 0) {
    console.warn('先生のメールが未登録のため、PRIVATE の写真を共有できませんでした: ' + file.getId());
    return;
  }
  // ウェブアプリは「デプロイした先生」として動くので、写真の持ち主はその先生。
  // 自分自身に閲覧権を渡そうとすると毎回失敗して警告が出るだけなので、外す。
  let owner = '';
  try { owner = normalizeEmail_(Session.getEffectiveUser().getEmail()); } catch (e) { owner = ''; }

  teachers.forEach(email => {
    if (owner && email === owner) return;
    try {
      file.addViewer(email);
    } catch (e) {
      console.warn('写真の閲覧権を渡せませんでした（宛先は Script Properties を参照）: ' + e);
    }
  });
}

/**
 * 画像の表示用URLを作る。
 * lh3.googleusercontent.com/d/<id> は「リンクを知っている全員」向けの
 * 公開CDN経路で、ドメイン限定・非公開のファイルでは表示できない。
 * drive.google.com/thumbnail は閲覧者の Google ログインで権限を見るため、
 * DOMAIN_WITH_LINK でも PRIVATE（＋先生に共有）でも表示できる。
 */
function buildImageUrl_(fileId) {
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1600';
}

// --- 記事関連 (Articles) ---

/**
 * 投稿された写真を、先生のドライブへ保存して表示用URLを返す。
 * 失敗しても記事そのものは残したいので、例外を外へ出さない（空文字を返す）。
 *
 * ⚠️ **ロックの外で呼ぶこと。** base64 の復号とドライブへの書き込みで
 *    1 件あたり数秒かかる。ロックの中に入れると 40 人ぶんが直列になり、
 *    合計が児童側の再送 3 回（2+4+6 秒）を軽く追い越す。
 */
function savePhoto_(id, base64, mimeType) {
  try {
    const props = PropertiesService.getScriptProperties();
    const folderId = props.getProperty(KEYS.FOLDER_ID);
    let folder = null;

    if (folderId) {
      try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = null; }
    }
    if (!folder) {
      folder = DriveApp.createFolder('新聞システム画像フォルダ');
      props.setProperty(KEYS.FOLDER_ID, folder.getId());
    }

    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, 'img_' + id);
    const file = folder.createFile(blob);
    applyPhotoSharing_(file);
    return buildImageUrl_(file.getId());
  } catch (e) {
    // 児童の本文やファイル名は出さない。記事IDとエラーだけ残す。
    console.error('画像保存エラー（記事 ' + id + '）: ' + e);
    return '';
  }
}

/**
 * 記事を保存する (Client -> Server)
 *
 * ロックで守るのは appendRow の 1 回だけにする（v5 §5-3）。
 * 写真の保存もシートをそろえる処理も、ロックの外で先に済ませておく。
 */
function saveArticle(data) {
  const id = Utilities.getUuid();
  const timestamp = new Date();

  // --- ここからロックの外 ---------------------------------------------
  const imageUrl = data.image ? savePhoto_(id, data.image, data.mimeType) : '';

  const ss = getDb_();                                  // 足りないシートと見出しをそろえる
  const sheet = ss.getSheetByName(SHEETS.ARTICLES);
  if (!sheet) throw new Error('記事シートを用意できませんでした。メニューの「4. シートを点検する」で確かめてください。');

  // 列の位置を決め打ちしない。見出し名で置き場所を決める。
  // 決め打ちにすると、先生が列を 1 本入れ替えただけで、本文が記者名の列に、
  // 記者名が投稿日時の列に入る。画面には何も出ないまま記事が壊れる。
  const cols = articleColumns_(sheet);
  const values = {
    'ID': id,
    'Title': data.title,
    'Body': data.body,
    'ImageURL': imageUrl,
    'Reporter': data.reporter,
    'Timestamp': timestamp,
    'Status': 'Pending',
    'Tag': data.tag || ''
  };
  const row = cols.header.map(name => {
    const k = schemaKey_(name);
    const hit = Object.keys(values).filter(n => schemaKey_(n) === k)[0];
    // アプリが知らない列（先生が足したメモ欄など）は空のままにする
    return hit === undefined ? '' : values[hit];
  });

  // --- ここからロックの中（1 行足すだけ） ------------------------------
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error('サーバーが混み合っています。もう一度送信ボタンを押してください。');
  }
  try {
    sheet.appendRow(row);
    return { success: true };
  } catch (e) {
    throw new Error('保存処理中にエラーが発生しました: ' + e);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 記事一覧を取得する (Server -> Admin Client)
 */
function getArticles() {
  requireTeacher_();
  try {
    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEETS.ARTICLES);
    if (!sheet) return [];

    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length < 2) return [];

    // 見出し名で列を決める。書き方の差（前後の空白・大文字小文字）は落として照らす。
    // 見つからない列は -1 にして「空」として扱う。
    // ★ ここで既定の位置に落とさないこと。見出しが無い列を 7 列目と決めつけると、
    //   先生が足したメモ欄をタグとして読み、記事一覧に関係のない文字列が並ぶ。
    //   「その列は無い」と正直に扱うほうが、点検メニューで気づける。
    const headers = values.shift().map(h => schemaKey_(h));
    const getIdx = (name) => headers.indexOf(schemaKey_(name));

    const idx = {
      id:       getIdx('ID'),
      title:    getIdx('Title'),
      body:     getIdx('Body'),
      img:      getIdx('ImageURL'),
      reporter: getIdx('Reporter'),
      ts:       getIdx('Timestamp'),
      tag:      getIdx('Tag')
    };

    /** その列が見つかっていて、その行にセルがあるときだけ中身を返す */
    const cell = (r, i) => (i >= 0 && i < r.length && r[i] !== null && r[i] !== undefined) ? r[i] : "";

    return values.reverse().map(r => {
      let ts = 0;
      const rawTs = cell(r, idx.ts);
      if (rawTs) {
        try { ts = new Date(rawTs).getTime(); } catch (e) { }
      }

      let rawImgUrl = String(cell(r, idx.img));
      if (rawImgUrl) {
        const idMatch = rawImgUrl.match(/id=([a-zA-Z0-9_-]+)/) || rawImgUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (idMatch && idMatch[1]) {
          // 過去に保存した lh3 形式（公開リンク前提）も、権限を見る形式へ読み替える
          rawImgUrl = buildImageUrl_(idMatch[1]);
        }
      }

      return {
        id:           String(cell(r, idx.id)),
        title:        String(cell(r, idx.title)),
        body:         String(cell(r, idx.body)),
        reporterName: String(cell(r, idx.reporter)),
        tag:          String(cell(r, idx.tag)),
        imageUrl:     rawImgUrl,
        timestamp:    ts
      };
    });
  } catch (e) {
    throw new Error("データ取得中にエラーが発生しました: " + e.toString());
  }
}

/**
 * 記事のタグを更新する
 */
function updateArticleTag(id, newTag) {
  requireTeacher_();

  // --- ロックの外 -------------------------------------------------------
  const sheet = getDb_().getSheetByName(SHEETS.ARTICLES);
  if (!sheet) return;

  // 列は見出し名で決める。無ければ articleColumns_ が右端に足す。
  // ★ 以前はここで ID を 0 列目・Tag を 7 列目に落としていた。
  //   Tag の列が無い（または並べ替えられている）シートでは、
  //   見出しの無い 8 列目や、まったく別の列にタグを書き込んでいた。
  // ★ articleColumns_ は自分でロックを取ることがあるので、**ロックの外**で呼ぶ。
  const cols = articleColumns_(sheet);
  const idCol = cols.col('ID');
  const tagCol = cols.col('Tag');
  if (!idCol || !tagCol) {
    throw new Error('記事シートに ID／Tag の列が見つかりません。メニューの「4. シートを点検する」で確かめてください。');
  }

  const data = sheet.getDataRange().getValues();
  let row = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol - 1]) === String(id)) { row = i + 1; break; }
  }
  if (!row) return;

  // --- ロックの中（1 セル書くだけ） --------------------------------------
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (e) {
    throw new Error('サーバーが混み合っています。もう一度お試しください。');
  }
  try {
    sheet.getRange(row, tagCol).setValue(newTag);
  } finally {
    lock.releaseLock();
  }
}

// --- 設定・保存関連 (System Data) ---

/**
 * 設定・テンプレートを入れる非表示シート。
 *
 * ★ 末尾の `_` は必須。以前は getSystemSheet という名前で、
 *   google.script.run から児童が直接呼べるトップレベル関数になっていた
 *   （戻り値は返せないが、呼ばれれば動く）。内部ヘルパーは必ず `_` を付ける。
 */
function getSystemSheet_() {
  const sheet = getDb_().getSheetByName(SHEETS.SYSTEM);
  if (!sheet) throw new Error('設定シートを用意できませんでした。');
  return sheet;
}

function saveLayoutState(name, json) {
  requireTeacher_();
  const sheet = getSystemSheet_();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
  sheet.appendRow(['LAYOUT', name, json, dateStr]);
  return { message: '✅ 保存しました' };
}

function getSavedList() {
  requireTeacher_();
  const sheet = getSystemSheet_();
  const rows = sheet.getDataRange().getValues();
  return rows
    .filter(r => r[0] === 'LAYOUT')
    .map(r => ({
      name: r[1],
      date: Utilities.formatDate(new Date(r[3]), Session.getScriptTimeZone(), "MM/dd HH:mm")
    }))
    .reverse();
}

function loadLayoutState(name) {
  requireTeacher_();
  const sheet = getSystemSheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0] === 'LAYOUT' && rows[i][1] === name) {
      return { success: true, data: rows[i][2] };
    }
  }
  return { success: false, message: 'データが見つかりません' };
}

function saveTemplate(name, json) {
  requireTeacher_();
  const sheet = getSystemSheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0] === 'TEMPLATE' && rows[i][1] === name) {
      sheet.deleteRow(i + 1);
    }
  }
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
  sheet.appendRow(['TEMPLATE', name, json, dateStr]);
  return { message: '✅ テンプレートを保存しました' };
}

function getTemplateList() {
  requireTeacher_();
  const sheet = getSystemSheet_();
  const rows = sheet.getDataRange().getValues();
  return rows.filter(r => r[0] === 'TEMPLATE').map(r => ({ name: r[1] })).reverse();
}

function loadTemplate(name) {
  requireTeacher_();
  const sheet = getSystemSheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0] === 'TEMPLATE' && rows[i][1] === name) {
      return { success: true, data: rows[i][2] };
    }
  }
  return { success: false, message: 'テンプレートが見つかりません' };
}

// --- タグ管理 (Tag Settings) ---
// 読み取り（getTagsSettings / getSchoolTags）は児童の投稿画面も使うので開いている。
// 書き換え（saveTagsSettings）は先生だけ。

function saveTagsSettings(tagsJson) {
  requireTeacher_();
  PropertiesService.getScriptProperties().setProperty(KEYS.TAG_SETTINGS, tagsJson);
  return { success: true };
}

function getTagsSettings() {
  const json = PropertiesService.getScriptProperties().getProperty(KEYS.TAG_SETTINGS);
  if (json) {
    return JSON.parse(json);
  } else {
    return DEFAULT_TAGS;
  }
}

function getSchoolTags() {
  const settings = getTagsSettings();
  return settings.map(t => t.name);
}
