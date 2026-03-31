/**
 * ヒアリングシート変換ツール用 Google Apps Script
 *
 * 【設置手順】
 * 1. 案件管理表スプレッドシートを開く
 * 2. 拡張機能 → Apps Script
 * 3. このコードを貼り付けて保存
 * 4. デプロイ → 新しいデプロイ → ウェブアプリ
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. デプロイURLをコピーしてHTMLツールに貼り付け
 */

// シート名の定義
var KISO_SHEET_NAME = '【基礎】案件管理表';
var MOKU_SHEET_NAME = '【目的地】案件管理表';
var HEADER_ROWS = 2; // ヘッダー行数

/**
 * POSTリクエストを処理
 */
function doPost(e) {
  try {
    var json = JSON.parse(e.postData.contents);
    var items = Array.isArray(json) ? json : [json];
    var results = [];

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var result = writeEntry(item);
      results.push(result);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, results: results }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GETリクエスト（接続テスト用）
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: 'ヒアリングシート変換ツール API ready' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 1件のエントリをスプレッドシートに書き込み
 */
function writeEntry(item) {
  var type = item.sheetType;
  var data = item.data;
  var sheetName;
  var fields;

  if (type === 'kiso') {
    sheetName = KISO_SHEET_NAME;
    fields = getKisoFields();
  } else if (type === 'mokutekichi') {
    sheetName = MOKU_SHEET_NAME;
    fields = getMokuFields();
  } else {
    throw new Error('不明なsheetType: ' + type);
  }

  // シートの名前を部分一致で検索（末尾スペース対策）
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = null;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf(sheetName) === 0 || sheetName.indexOf(sheets[i].getName()) === 0) {
      sheet = sheets[i];
      break;
    }
  }
  if (!sheet) {
    // 完全一致で再試行
    sheet = ss.getSheetByName(sheetName);
  }
  if (!sheet) {
    throw new Error('シート "' + sheetName + '" が見つかりません');
  }

  // 最終行とNo.を取得
  var lastRow = sheet.getLastRow();
  var nextRow = lastRow + 1;
  var maxNo = 0;

  if (lastRow > HEADER_ROWS) {
    var noRange = sheet.getRange(HEADER_ROWS + 1, 1, lastRow - HEADER_ROWS, 1);
    var noValues = noRange.getValues();
    for (var i = 0; i < noValues.length; i++) {
      var v = noValues[i][0];
      if (typeof v === 'number' && v > maxNo) maxNo = v;
    }
  }
  var nextNo = maxNo + 1;

  // A列にNo.を書き込み
  sheet.getRange(nextRow, 1).setValue(nextNo);

  // 各フィールドを書き込み（E列=5列目, F=6, ... V=22）
  for (var i = 0; i < fields.length; i++) {
    var f = fields[i];
    var value = data[f.key];
    if (value !== undefined && value !== null && value !== '') {
      sheet.getRange(nextRow, f.colNum).setValue(value);
    }
  }

  return {
    name: data.propertyName || data.facilityName || '不明',
    sheet: sheet.getName(),
    row: nextRow,
    no: nextNo
  };
}

/**
 * 基礎シートのフィールド定義 (E〜V列)
 */
function getKisoFields() {
  return [
    { key: 'leadPartner',      colNum: 5 },  // E
    { key: 'salesPerson',       colNum: 6 },  // F
    { key: 'applicationPeriod', colNum: 7 },  // G
    { key: 'propertyName',      colNum: 8 },  // H
    { key: 'plan',              colNum: 9 },  // I
    { key: 'mansionType',       colNum: 10 }, // J
    { key: 'kwNumber',          colNum: 11 }, // K
    { key: 'newOrExisting',     colNum: 12 }, // L
    { key: 'address',           colNum: 13 }, // M
    { key: 'landOwner',         colNum: 14 }, // N
    { key: 'units',             colNum: 15 }, // O
    { key: 'parkingCount',      colNum: 16 }, // P
    { key: 'applicationCount',  colNum: 17 }, // Q
    { key: 'companyName',       colNum: 18 }, // R
    { key: 'companyAddress',    colNum: 19 }, // S
    { key: 'csCompany',         colNum: 20 }, // T
    { key: 'csPersonName',      colNum: 21 }, // U
    { key: 'csEmail',           colNum: 22 }, // V
  ];
}

/**
 * 目的地シートのフィールド定義 (E〜T列)
 */
function getMokuFields() {
  return [
    { key: 'leadPartner',      colNum: 5 },  // E
    { key: 'salesPerson',       colNum: 6 },  // F
    { key: 'applicationPeriod', colNum: 7 },  // G
    { key: 'facilityName',      colNum: 8 },  // H
    { key: 'industry',          colNum: 9 },  // I
    { key: 'plan',              colNum: 10 }, // J
    { key: 'kwNumber',          colNum: 11 }, // K
    { key: 'address',           colNum: 12 }, // L
    { key: 'landOwner',         colNum: 13 }, // M
    { key: 'parkingCount',      colNum: 14 }, // N
    { key: 'applicationCount',  colNum: 15 }, // O
    { key: 'companyName',       colNum: 16 }, // P
    { key: 'companyAddress',    colNum: 17 }, // Q
    { key: 'csCompany',         colNum: 18 }, // R
    { key: 'csPersonName',      colNum: 19 }, // S
    { key: 'csEmail',           colNum: 20 }, // T
  ];
}
