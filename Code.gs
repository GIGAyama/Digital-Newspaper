// ==================================================
// 📰 デジタル・クラス新聞社 システムコード (Ver.7.1)
// ==================================================

// ★★★ ここで「基本のタグ」を設定できます ★★★
const DEFAULT_TAGS = ['学校生活', '行事', '学習', '委員会', 'クラブ', '休み時間', 'その他'];

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
  const currentId = PropertiesService.getScriptProperties().getProperty('IMAGE_FOLDER_ID') || '';
  const result = ui.prompt('写真保存用フォルダの設定', 'GoogleドライブのフォルダIDを入力してください：\n(現在: ' + (currentId ? currentId : '未設定') + ')', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() == ui.Button.OK) {
    const newId = result.getResponseText().trim();
    if (!newId) { ui.alert('❌ IDが空です'); return; }
    PropertiesService.getScriptProperties().setProperty('IMAGE_FOLDER_ID', newId);
    ui.alert('✅ 設定を保存しました！');
  }
}

function showAdminUrl() {
  const ui = SpreadsheetApp.getUi();
  let url = ScriptApp.getService().getUrl();
  if (!url) { ui.alert('⚠️ まず「デプロイ」を実行して、WebアプリのURLを発行してください。'); return; }
  
  const htmlOutput = HtmlService.createHtmlOutput(
    '<div style="text-align:center; padding:20px; font-family:sans-serif;">' +
    '<p>以下のリンクから新聞編集室へ移動します。</p>' +
    '<a href="' + url + '?p=admin" target="_blank" style="background:#007bff;color:white;padding:12px 25px;text-decoration:none;border-radius:5px;font-weight:bold;display:inline-block;box-shadow:0 2px 5px rgba(0,0,0,0.2);">🚀 編集室に入る</a>' +
    '<p style="margin-top:15px; font-size:0.85rem; color:#666;">※ ポップアップブロックされた場合は許可してください</p>' +
    '</div>'
  ).setWidth(400).setHeight(200);
  ui.showModalDialog(htmlOutput, '管理画面へのアクセス');
}

function doGet(e) {
  const folderId = PropertiesService.getScriptProperties().getProperty('IMAGE_FOLDER_ID');
  if (!folderId) return HtmlService.createHtmlOutput('<div style="padding:20px; text-align:center; color:red;"><h3>⚠️ エラー</h3><p>写真フォルダIDが設定されていません。<br>スプレッドシートのメニュー「📰 新聞システム」から設定を行ってください。</p></div>');

  let page = e.parameter.p || 'index';
  if (!['index', 'admin'].includes(page)) page = 'index';

  const template = HtmlService.createTemplateFromFile(page);
  template.appUrl = ScriptApp.getService().getUrl();
  
  return template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle(page === 'admin' ? '📰 新聞編集室' : '📮 記者投稿ポスト')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- 児童用: タグリストを取得する関数 (New) ---
function getSchoolTags() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Articles');
  // デフォルトタグをセット（重複排除のためSetを使用）
  let tags = new Set(DEFAULT_TAGS); 
  
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // H列 (8列目) のタグを取得し、過去に使われたタグも選択肢に加える
      const data = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
      data.forEach(row => {
        if (row[0]) tags.add(row[0]);
      });
    }
  }
  // 配列に戻してソートして返す
  return Array.from(tags).sort();
}

function saveArticle(formObject) {
  try {
    const folderId = PropertiesService.getScriptProperties().getProperty('IMAGE_FOLDER_ID');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Articles');
    if(!sheet) throw new Error('Articlesシートが見つかりません');

    const folder = DriveApp.getFolderById(folderId);
    
    let imageUrl = '';
    if (formObject.imageFile && formObject.imageFile.length > 0) {
      const blob = formObject.imageFile;
      const fileName = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '_' + formObject.reporterName;
      const file = folder.createFile(blob).setName(fileName);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    }

    sheet.appendRow([
      Utilities.getUuid(),
      formObject.title,
      formObject.body,
      imageUrl,
      formObject.reporterName,
      new Date(),
      'Pending',
      formObject.tag || '' // Tag (修正: フォームからタグを受け取る)
    ]);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getArticles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Articles');
  if(!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  
  return data.map((row) => {
    let imgUrl = row[3];
    if (imgUrl) {
       const idMatch = imgUrl.match(/id=([a-zA-Z0-9_-]+)/) || imgUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
       if (idMatch) imgUrl = "https://lh3.googleusercontent.com/d/" + idMatch[1];
    }
    return {
      id: row[0],
      title: row[1],
      body: row[2],
      imageUrl: imgUrl,
      reporterName: row[4],
      date: Utilities.formatDate(new Date(row[5]), 'Asia/Tokyo', 'MM/dd HH:mm'),
      timestamp: new Date(row[5]).getTime(),
      tag: row[7] || ''
    };
  }).reverse();
}

// --- タグ更新機能 ---
function updateArticleTag(articleId, newTag) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Articles');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == articleId) {
      sheet.getRange(i + 1, 8).setValue(newTag);
      return { success: true };
    }
  }
  return { success: false };
}

// --- 編集状態保存 ---
function getSystemSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(['Name', 'JsonData', 'UpdatedAt']);
    sheet.hideSheet();
  }
  return sheet;
}

function saveLayoutState(name, jsonData) {
  const sheet = getSystemSheet('SystemData');
  const data = sheet.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) { row = i + 1; break; }
  }
  if (row > 0) {
    sheet.getRange(row, 2).setValue(jsonData);
    sheet.getRange(row, 3).setValue(new Date());
  } else {
    sheet.appendRow([name, jsonData, new Date()]);
  }
  return { success: true, message: '✅ 保存しました！' };
}

function getSavedList() {
  const sheet = getSystemSheet('SystemData');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return data.map(row => ({ 
    name: row[0], 
    date: Utilities.formatDate(new Date(row[2]), 'Asia/Tokyo', 'MM/dd HH:mm') 
  })).reverse();
}

function loadLayoutState(name) {
  const sheet = getSystemSheet('SystemData');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) return { success: true, data: data[i][1] };
  }
  return { success: false, message: 'データが見つかりません' };
}

// --- テンプレート機能 ---
function saveTemplate(name, jsonData) {
  const sheet = getSystemSheet('Templates');
  const data = sheet.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) { row = i + 1; break; }
  }
  if (row > 0) {
    sheet.getRange(row, 2).setValue(jsonData);
    sheet.getRange(row, 3).setValue(new Date());
  } else {
    sheet.appendRow([name, jsonData, new Date()]);
  }
  return { success: true, message: '✅ テンプレート「' + name + '」を登録しました！' };
}

function getTemplateList() {
  const sheet = getSystemSheet('Templates');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return data.map(row => ({ name: row[0] })).reverse();
}

function loadTemplate(name) {
  const sheet = getSystemSheet('Templates');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) return { success: true, data: data[i][1] };
  }
  return { success: false, message: 'テンプレートが見つかりません' };
}
