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
  FOLDER_ID: 'IMAGE_FOLDER_ID', // 画像保存先フォルダID
  TAG_SETTINGS: 'TAG_SETTINGS'  // タグ設定JSON
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
// 2. スプレッドシート連携・メニュー (Spreadsheet UI)
// ==================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📰 新聞システム')
    .addItem('1. 写真フォルダの設定', 'setFolderId')
    .addSeparator()
    .addItem('2. 先生用管理画面を開く', 'showAdminUrl')
    .addToUi();
}

function setFolderId() {
  const ui = SpreadsheetApp.getUi();
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
  const page = e.parameter.p;
  let template;
  let title;

  if (page === 'admin') {
    template = HtmlService.createTemplateFromFile('admin');
    title = 'デジタル新聞編集室';
  } else {
    template = HtmlService.createTemplateFromFile('index');
    title = 'デジタルクラス新聞社';
  }

  return template.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=1teNY1fFaXzZC3HivQIQ4t9nu49xYsbJN&.png');
}

// ==================================================
// 4. データ処理ロジック (Data Logic)
// ==================================================

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
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        imageUrl = "https://lh3.googleusercontent.com/d/" + file.getId();

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
          rawImgUrl = "https://lh3.googleusercontent.com/d/" + idMatch[1];
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
  const sheet = getSystemSheet();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
  sheet.appendRow(['LAYOUT', name, json, dateStr]);
  return { message: '✅ 保存しました' };
}

function getSavedList() {
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
  const sheet = getSystemSheet();
  const rows = sheet.getDataRange().getValues();
  return rows.filter(r => r[0] === 'TEMPLATE').map(r => ({ name: r[1] })).reverse();
}

function loadTemplate(name) {
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

function saveTagsSettings(tagsJson) {
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
