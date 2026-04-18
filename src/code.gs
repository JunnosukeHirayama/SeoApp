/**
 * SEO Manager
 */

const SHEET_NAMES = {
  ARTICLES: 'Articles',
  LINKS: 'Links',
  CLUSTERS: 'Clusters',
  MEMBERS: 'ClusterMembers'
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('SEO Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('SEO管理')
    .addItem('ID修復ツール', 'repairIds')
    .addToUi();
}

// --- データ取得 ---
function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {};
  try {
    for (const key in SHEET_NAMES) {
      sheets[key] = ss.getSheetByName(SHEET_NAMES[key]);
      if (!sheets[key]) throw new Error(`シート「${SHEET_NAMES[key]}」が見つかりません。`);
    }
  } catch (e) {
    throw new Error(e.message);
  }

  return {
    articles: fetchSheetData(sheets.ARTICLES, 3),
    links: fetchSheetData(sheets.LINKS, 2),
    clusters: fetchSheetData(sheets.CLUSTERS, 3),
    members: fetchSheetData(sheets.MEMBERS, 2)
  };
}

function fetchSheetData(sheet, colCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, colCount).getDisplayValues();
  return data
    .filter(row => row[0] && row[0].trim() !== "")
    .map(row => row.map(cell => String(cell).trim()));
}

// --- 記事管理アクション ---

// 追加
function addArticle(title, url) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ARTICLES);
  sheet.appendRow([Utilities.getUuid(), title, url]);
  return "記事を追加しました";
}

// 更新 (新規追加)
function updateArticle(id, title, url) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.ARTICLES);
  const targetId = String(id).trim();
  
  if (sheet.getLastRow() < 2) throw new Error("記事が見つかりません");
  

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  const rowIndex = data.findIndex(row => String(row[0]).trim() === targetId);
  
  if (rowIndex === -1) throw new Error("対象の記事が見つかりませんでした");
  
  
  sheet.getRange(rowIndex + 2, 2, 1, 2).setValues([[title, url]]);
  
  return "記事情報を更新しました";
}

// 削除
function deleteArticle(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetId = String(id).trim();
  
 
  const aSheet = ss.getSheetByName(SHEET_NAMES.ARTICLES);
  deleteRowById(aSheet, targetId, 1); // 1列目がID

  
  const lSheet = ss.getSheetByName(SHEET_NAMES.LINKS);
  filterOutRows(lSheet, row => String(row[0]).trim() !== targetId && String(row[1]).trim() !== targetId);

  
  const mSheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
  filterOutRows(mSheet, row => String(row[1]).trim() !== targetId); // 2列目が記事ID

 
  const cSheet = ss.getSheetByName(SHEET_NAMES.CLUSTERS);
  if (cSheet.getLastRow() >= 2) {
    const cData = cSheet.getRange(2, 1, cSheet.getLastRow() - 1, 3).getDisplayValues();
    const updatedCData = cData.map(row => {
      // PillarID(3列目)が削除対象なら空文字にする
      if (String(row[2]).trim() === targetId) return [row[0], row[1], ""]; 
      return row;
    });
    cSheet.getRange(2, 1, updatedCData.length, 3).setValues(updatedCData);
  }

  SpreadsheetApp.flush();
  return "記事を削除しました (関連リンクも整理しました)";
}

// --- ヘルパー関数 ---


function filterOutRows(sheet, keepConditionFn) {
  if (sheet.getLastRow() < 2) return;
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const data = range.getDisplayValues();
  const newData = data.filter(keepConditionFn);
  
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getLastColumn()).clearContent();
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
  }
}


function deleteRowById(sheet, id, idColIndex) {
  if (sheet.getLastRow() < 2) return;
  const data = sheet.getRange(2, idColIndex, sheet.getLastRow() - 1, 1).getDisplayValues();

  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim() === id) {
      sheet.deleteRow(i + 2);
    }
  }
}

// --- 基本機能 ---

function saveInternalLinks(sourceId, targetIds) {
  if (!sourceId || String(sourceId).trim() === "") throw new Error("IDエラー");
  const sId = String(sourceId).trim();
  const tIds = targetIds.map(id => String(id).trim());
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LINKS);
  sheet.getRange("A:B").setNumberFormat("@");
  
  let allRows = [];
  if (sheet.getLastRow() >= 2) allRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
  const keptRows = allRows.filter(row => String(row[0]).trim() !== sId);
  const newRows = tIds.map(tid => [sId, tid]);
  const finalRows = [...keptRows, ...newRows];
  
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) sheet.getRange(2, 1, maxRows - 1, 2).clearContent();
  if (finalRows.length > 0) sheet.getRange(2, 1, finalRows.length, 2).setValues(finalRows);
  
  SpreadsheetApp.flush();
  return `保存完了: ${tIds.length}件`;
}

function repairIds() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ARTICLES);
  if (sheet.getLastRow() < 2) return;
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  const values = range.getValues();
  let count = 0;
  const newValues = values.map(r => {
    if (!r[0] || String(r[0]).trim() === "") { count++; return [Utilities.getUuid()]; }
    return [r[0]];
  });
  if(count>0) { range.setValues(newValues); SpreadsheetApp.getUi().alert(`${count}件修復`); }
  else { SpreadsheetApp.getUi().alert("異常なし"); }
}

function saveCluster(cId, name, pId, mIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cSheet = ss.getSheetByName(SHEET_NAMES.CLUSTERS);
  const mSheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
  const clusterId = cId ? String(cId).trim() : Utilities.getUuid();
  
  let cRows = [];
  if(cSheet.getLastRow() >= 2) cRows = cSheet.getRange(2, 1, cSheet.getLastRow()-1, 3).getDisplayValues();
  const idx = cRows.findIndex(r => String(r[0]).trim() === clusterId);
  if(idx >= 0) cRows[idx] = [clusterId, name, pId];
  else cRows.push([clusterId, name, pId]);
  cSheet.getRange('A2:C').clearContent();
  if(cRows.length > 0) cSheet.getRange(2, 1, cRows.length, 3).setValues(cRows);

  let mRows = [];
  if(mSheet.getLastRow() >= 2) mRows = mSheet.getRange(2, 1, mSheet.getLastRow()-1, 2).getDisplayValues();
  const kept = mRows.filter(r => String(r[0]).trim() !== clusterId);
  const newM = mIds.map(mid => [clusterId, String(mid).trim()]);
  const finalM = [...kept, ...newM];
  mSheet.getRange('A2:B').clearContent();
  if(finalM.length > 0) mSheet.getRange(2, 1, finalM.length, 2).setValues(finalM);
  
  return "保存しました";
}

function deleteCluster(cId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = String(cId).trim();
  const cSheet = ss.getSheetByName(SHEET_NAMES.CLUSTERS);
  const mSheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
  filterOutRows(cSheet, r => String(r[0]).trim() !== id);
  filterOutRows(mSheet, r => String(r[0]).trim() !== id);
  return "削除しました";
}