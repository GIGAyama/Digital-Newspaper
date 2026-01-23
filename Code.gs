// ==================================================
// 📰 デジタル・クラス新聞社 システムコード (Ver.5.0 Fixed)
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
  const currentId = PropertiesService.getScriptProperties().getProperty('IMAGE_FOLDER_ID') || '';
  const result = ui.prompt('写真保存用フォルダの設定', 'フォルダIDを入力：\n(現在: ' + (currentId ? currentId : '未設定') + ')', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('IMAGE_FOLDER_ID', result.getResponseText().trim());
    ui.alert('✅ 設定を保存しました！');
  }
}

function showAdminUrl() {
  const ui = SpreadsheetApp.getUi();
  let url = ScriptApp.getService().getUrl();
  if (!url) { ui.alert('⚠️ まずデプロイしてください。'); return; }
  const htmlOutput = HtmlService.createHtmlOutput(
    '<div style="text-align:center; padding:20px; font-family:sans-serif;">' +
    '<p>以下のリンクから管理画面へ移動します。</p>' +
    '<a href="' + url + '?p=admin" target="_top" style="background:#007bff;color:white;padding:10px 20px;text-decoration:none;border-radius:5px;font-weight:bold;">🚀 編集画面へ</a>' +
    '</div>'
  ).setWidth(350).setHeight(150);
  ui.showModalDialog(htmlOutput, '管理画面');
}

function doGet(e) {
  const folderId = PropertiesService.getScriptProperties().getProperty('IMAGE_FOLDER_ID');
  if (!folderId) return HtmlService.createHtmlOutput('⚠️ エラー：写真フォルダIDが設定されていません。スプレッドシートのメニューから設定してください。');

  let page = e.parameter.p || 'index';
  if (!['index', 'admin'].includes(page)) page = 'index';

  const template = HtmlService.createTemplateFromFile(page);
  template.appUrl = ScriptApp.getService().getUrl(); 
  return template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle(page === 'admin' ? '新聞編集室' : '記事投稿ポスト');
}

function saveArticle(formObject) {
  try {
    const folderId = PropertiesService.getScriptProperties().getProperty('IMAGE_FOLDER_ID');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
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
      ''
    ]);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getArticles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  return data.map((row, index) => {
    let imgUrl = row[3];
    if (imgUrl && imgUrl.indexOf('drive.google.com') !== -1) {
       let idMatch = imgUrl.match(/id=([a-zA-Z0-9_-]+)/) || imgUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
       if (idMatch) imgUrl = "https://lh3.googleusercontent.com/d/" + idMatch[1];
    }
    return {
      id: row[0],
      title: row[1],
      body: row[2],
      imageUrl: imgUrl,
      reporterName: row[4],
      date: Utilities.formatDate(new Date(row[5]), 'Asia/Tokyo', 'MM/dd HH:mm'),
      timestamp: new Date(row[5]).getTime()
    };
  }).reverse();
}

// --- 保存機能 ---
function getSystemSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('SystemData');
  if (!sheet) {
    sheet = ss.insertSheet('SystemData');
    sheet.appendRow(['SaveName', 'JsonData', 'UpdatedAt']);
    sheet.hideSheet();
  }
  return sheet;
}

function saveLayoutState(name, jsonData) {
  const sheet = getSystemSheet();
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
  return { success: true, message: '✅ 保存しました: ' + name };
}

function getSavedList() {
  const sheet = getSystemSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  // 新しい順に
  return data.map(row => ({ 
    name: row[0], 
    date: Utilities.formatDate(new Date(row[2]), 'Asia/Tokyo', 'MM/dd HH:mm') 
  })).reverse();
}

function loadLayoutState(name) {
  const sheet = getSystemSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      return { success: true, data: data[i][1] };
    }
  }
  return { success: false, message: 'データが見つかりません' };
}
