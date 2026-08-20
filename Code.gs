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
// 1.5 教員判定 (Authorization)
// ==================================================
//
// 管理画面は URL を知っていれば誰でも開けてしまう作りだった。
// 画面を隠すだけでは足りない（google.script.run から関数を直接呼べる）ので、
// 「画面の入口」と「管理系のサーバー関数」の両方で同じ判定を通す。

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
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  ensureOwnerEmail_();
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

// ==================================================
// 2. スプレッドシート連携・メニュー (Spreadsheet UI)
// ==================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📰 新聞システム')
    .addItem('1. 写真フォルダの設定', 'setFolderId')
    .addItem('2. 先生のメールアドレスを設定', 'setTeacherEmails')
    .addSeparator()
    .addItem('3. 先生用管理画面を開く', 'showAdminUrl')
    .addToUi();
}

function setFolderId() {
  const ui = SpreadsheetApp.getUi();
  // 初回セットアップ実行者を管理者として記録する（TEACHER_EMAILS 未設定時のみ）
  ensureOwnerEmail_();
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

  if (!isTeacher_()) {
    ui.alert(
      '⚠️ このアカウントは管理画面を使えません。\n' +
      'メニューの「2. 先生のメールアドレスを設定」で、いま使っているアカウントを登録してください。'
    );
    return;
  }

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
    template = HtmlService.createTemplateFromFile('index');
    title = 'デジタルクラス新聞社';
  }
  // 児童が見る画面に編集室への入口を出さないための目印（index.html で使う）
  template.isTeacher = teacher;

  return template.evaluate()
    .setTitle(title)
    // viewport-fit=cover が無いと、切り欠きのある端末で env(safe-area-inset-*) が使えない。
    // GAS は画面を iframe で包むため、HTML 側の <meta> だけでは足りず、
    // サーバー側で足すこのタグにも要る（v5 §5）。前回 index.html だけを直して漏れていた。
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
  teachers.forEach(email => {
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
 * 記事を保存する (Client -> Server)
 * ★ ロック処理を追加し、同時書き込み時のデータ破損を防ぎます
 */
function saveArticle(data) {
  // 排他制御ロックを取得 (最大10秒待機)
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // ロック獲得待ち
  } catch (e) {
    throw new Error("サーバーが混み合っています。もう一度送信ボタンを押してください。");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEETS.ARTICLES);

    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.ARTICLES);
      // 旧構成互換: [ID, Title, Body, ImageURL, Reporter, Timestamp, Status, Tag]
      sheet.appendRow(['ID', 'Title', 'Body', 'ImageURL', 'Reporter', 'Timestamp', 'Status', 'Tag']);
    }

    const id = Utilities.getUuid();
    const timestamp = new Date();
    let imageUrl = '';

    // 画像処理
    if (data.image) {
      try {
        const props = PropertiesService.getScriptProperties();
        let folderId = props.getProperty(KEYS.FOLDER_ID);
        let folder;

        if (folderId) {
          try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = null; }
        }

        if (!folder) {
          folder = DriveApp.createFolder("新聞システム画像フォルダ");
          props.setProperty(KEYS.FOLDER_ID, folder.getId());
        }

        const blob = Utilities.newBlob(Utilities.base64Decode(data.image), data.mimeType, "img_" + id);
        const file = folder.createFile(blob);
        applyPhotoSharing_(file);
        imageUrl = buildImageUrl_(file.getId());

      } catch (e) {
        console.error("画像保存エラー: " + e.toString());
      }
    }

    // スプレッドシートへの書き込み
    sheet.appendRow([
      id, 
      data.title, 
      data.body, 
      imageUrl, 
      data.reporter, 
      timestamp, 
      'Pending', 
      data.tag || ''
    ]);
    
    return { success: true };

  } catch (e) {
    // 予期せぬエラー
    throw new Error("保存処理中にエラーが発生しました: " + e.toString());
  } finally {
    // 処理終了後に必ずロックを解除
    lock.releaseLock();
  }
}

/**
 * 記事一覧を取得する (Server -> Admin Client)
 */
function getArticles() {
  requireTeacher_();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.ARTICLES);
    if (!sheet) return [];

    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length < 2) return [];

    const headers = values.shift();
    const getIdx = (name) => headers.indexOf(name);
    
    const idx = {
      id:       getIdx('ID')       !== -1 ? getIdx('ID')       : 0,
      title:    getIdx('Title')    !== -1 ? getIdx('Title')    : 1,
      body:     getIdx('Body')     !== -1 ? getIdx('Body')     : 2,
      img:      getIdx('ImageURL') !== -1 ? getIdx('ImageURL') : 3,
      reporter: getIdx('Reporter') !== -1 ? getIdx('Reporter') : 4,
      ts:       getIdx('Timestamp')!== -1 ? getIdx('Timestamp'): 5,
      tag:      getIdx('Tag')      !== -1 ? getIdx('Tag')      : 7
    };

    return values.reverse().map(r => {
      let ts = 0;
      if (idx.ts < r.length && r[idx.ts]) {
        try { ts = new Date(r[idx.ts]).getTime(); } catch (e) { }
      }

      let rawImgUrl = (idx.img < r.length) ? String(r[idx.img]) : "";
      if (rawImgUrl) {
        const idMatch = rawImgUrl.match(/id=([a-zA-Z0-9_-]+)/) || rawImgUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (idMatch && idMatch[1]) {
          // 過去に保存した lh3 形式（公開リンク前提）も、権限を見る形式へ読み替える
          rawImgUrl = buildImageUrl_(idMatch[1]);
        }
      }

      return {
        id:           (idx.id < r.length)       ? String(r[idx.id]) : "",
        title:        (idx.title < r.length)    ? String(r[idx.title]) : "",
        body:         (idx.body < r.length)     ? String(r[idx.body]) : "",
        reporterName: (idx.reporter < r.length) ? String(r[idx.reporter]) : "",
        tag:          (idx.tag < r.length)      ? String(r[idx.tag]) : "",
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
  // ロック取得 (短時間の書き込み)
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.ARTICLES);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    let idCol = data[0].indexOf('ID');
    let tagCol = data[0].indexOf('Tag');
    
    if (idCol === -1) idCol = 0;
    if (tagCol === -1) tagCol = 7;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        sheet.getRange(i + 1, tagCol + 1).setValue(newTag);
        break;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// --- 設定・保存関連 (System Data) ---

function getSystemSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.SYSTEM);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.SYSTEM);
    sheet.appendRow(['Type', 'Name', 'Data', 'Date']);
    sheet.hideSheet();
  }
  return sheet;
}

function saveLayoutState(name, json) {
  requireTeacher_();
  const sheet = getSystemSheet();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
  sheet.appendRow(['LAYOUT', name, json, dateStr]);
  return { message: '✅ 保存しました' };
}

function getSavedList() {
  requireTeacher_();
  const sheet = getSystemSheet();
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
  const sheet = getSystemSheet();
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
  const sheet = getSystemSheet();
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
  const sheet = getSystemSheet();
  const rows = sheet.getDataRange().getValues();
  return rows.filter(r => r[0] === 'TEMPLATE').map(r => ({ name: r[1] })).reverse();
}

function loadTemplate(name) {
  requireTeacher_();
  const sheet = getSystemSheet();
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
