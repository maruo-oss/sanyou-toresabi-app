/**
 * かんばん照合アプリ - メインサーバーサイド処理
 *
 * 【前提】
 * このスクリプトは「照合データ」スプレッドシートに直接バインド（コンテナバインド）して使います。
 * スプレッドシートの「拡張機能」→「Apps Script」から開いてコードを貼り付けてください。
 *
 * 【スクリプトプロパティの設定】
 * GASエディタ → プロジェクトの設定（⚙） → スクリプトプロパティ に以下を追加：
 *   EMPLOYEE_SPREADSHEET_ID : 社員マスタ用スプレッドシートのID
 *   EMPLOYEE_SHEET_NAME     : 社員マスタのシート名（省略時: "社員マスタ"）
 *
 * ※スプレッドシートIDはURLの /d/ と /edit の間の文字列です
 *   例: https://docs.google.com/spreadsheets/d/【ここがID】/edit
 */

// ==== 設定 ====
var SHEET_DATA = '照合データ';

// 照合データ用: バインドされたスプレッドシート自身を返す
function getDataSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// 社員マスタ用: スクリプトプロパティから別スプレッドシートを参照
function getEmployeeConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    spreadsheetId: props.getProperty('EMPLOYEE_SPREADSHEET_ID'),
    sheetName: props.getProperty('EMPLOYEE_SHEET_NAME') || '社員マスタ'
  };
}

// ==== Webアプリエントリポイント ====
function doGet(e) {
  var page = e.parameter.page || 'index';
  var template;

  switch (page) {
    case 'scanner':
      template = HtmlService.createTemplateFromFile('scanner');
      break;
    case 'list':
      template = HtmlService.createTemplateFromFile('list');
      break;
    default:
      template = HtmlService.createTemplateFromFile('index');
  }

  return template.evaluate()
    .setTitle('かんばん照合アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// HTMLインクルード用
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// WebアプリのベースURL取得
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// ==== 社員マスタ（別スプレッドシートから取得） ====
function getEmployeeName(employeeCode) {
  var config = getEmployeeConfig();
  var ss = SpreadsheetApp.openById(config.spreadsheetId);
  var sheet = ss.getSheetByName(config.sheetName);
  var data = sheet.getDataRange().getValues();

  // A列: 社員番号, B列: 氏名, C列: 工場, N列(index 13): 退社フラグ
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(employeeCode)) {
      var resigned = data[i][13];
      if (resigned === true || String(resigned).toLowerCase() === 'true') {
        return { success: false, name: '', factory: '', resigned: true };
      }
      return { success: true, name: data[i][1], factory: data[i][2] || '' };
    }
  }
  return { success: false, name: '', factory: '' };
}

// ==== ヘッダー行の確認・作成 ====
var HEADERS = [
  'タイムスタンプ', '社員番号', '出荷日', '出荷便', '背番号',
  '製品型番', 'eかんばん枝番', '切断仕上No.①', '切断仕上No.②', '切断仕上No.③'
];

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    // ヘッダー行の書式設定
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a2a5e');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}

// ==== 照合データ書き込み ====
function saveRecord(record) {
  try {
    var ss = getDataSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DATA);
    ensureHeaders(sheet);

    // 二重登録チェック: 同じeかんばん枝番 + 出荷日 + 出荷便の組み合わせ
    var existing = sheet.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (String(existing[i][6]) === String(record.kanbanEdaban) &&
          String(existing[i][2]) === String(record.shippingDate) &&
          String(existing[i][3]) === String(record.shippingBin)) {
        return { success: false, message: 'このeかんばんは既に登録済みです（枝番: ' + record.kanbanEdaban + '）' };
      }
    }

    // 新規レコード追加
    sheet.appendRow([
      new Date(),                    // A: タイムスタンプ
      record.employeeCode,           // B: 社員番号
      record.shippingDate,           // C: 出荷日
      record.shippingBin,            // D: 出荷便
      record.sebangoNo,              // E: 背番号
      record.productModel,           // F: 製品型番
      record.kanbanEdaban,           // G: eかんばん枝番
      record.setsudanNo1 || '',      // H: 切断仕上No.①
      record.setsudanNo2 || '',      // I: 切断仕上No.②
      record.setsudanNo3 || ''       // J: 切断仕上No.③
    ]);

    return { success: true, message: '保存しました' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

// ==== 照合データ取得（集計画面用） ====
function getRecords(shippingDate, shippingBin) {
  try {
    var ss = getDataSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DATA);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var records = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowDate = Utilities.formatDate(new Date(row[2]), 'Asia/Tokyo', 'yyyy-MM-dd');
      var matchDate = !shippingDate || rowDate === shippingDate;
      var matchBin = !shippingBin || String(row[3]) === String(shippingBin);

      if (matchDate && matchBin) {
        records.push({
          timestamp: Utilities.formatDate(new Date(row[0]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
          employeeCode: row[1],
          shippingDate: rowDate,
          shippingBin: row[3],
          sebangoNo: row[4],
          productModel: row[5],
          kanbanEdaban: row[6],
          setsudanNo1: row[7],
          setsudanNo2: row[8],
          setsudanNo3: row[9]
        });
      }
    }

    return { success: true, records: records };
  } catch (e) {
    return { success: false, records: [], message: 'エラー: ' + e.message };
  }
}

// ==== 集計データ取得（背番号ごとのケース数） ====
function getSummary(shippingDate, shippingBin) {
  try {
    var result = getRecords(shippingDate, shippingBin);
    if (!result.success) return result;

    var summary = {};
    result.records.forEach(function(rec) {
      var key = rec.sebangoNo;
      if (!summary[key]) {
        summary[key] = { sebangoNo: key, productModel: rec.productModel, count: 0 };
      }
      summary[key].count++;
    });

    // オブジェクトを配列に変換し、背番号順にソート
    var summaryArray = Object.keys(summary).map(function(key) {
      return summary[key];
    }).sort(function(a, b) {
      return Number(a.sebangoNo) - Number(b.sebangoNo);
    });

    return { success: true, summary: summaryArray, totalCount: result.records.length };
  } catch (e) {
    return { success: false, summary: [], message: 'エラー: ' + e.message };
  }
}

// ==== レコード削除（取消用） ====
function deleteRecord(timestamp, kanbanEdaban) {
  try {
    var ss = getDataSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DATA);
    var data = sheet.getDataRange().getValues();

    for (var i = data.length - 1; i >= 1; i--) {
      var rowTimestamp = Utilities.formatDate(new Date(data[i][0]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
      if (rowTimestamp === timestamp && String(data[i][6]) === String(kanbanEdaban)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: '削除しました' };
      }
    }
    return { success: false, message: '該当レコードが見つかりません' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

// ==== 既存データにヘッダーを追加（1回だけ手動実行） ====
function addHeadersToExistingSheet() {
  var ss = getDataSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_DATA);
  // 1行目にヘッダーを挿入
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, HEADERS.length).setBackground('#1a2a5e');
  sheet.getRange(1, 1, 1, HEADERS.length).setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  Logger.log('ヘッダーを追加しました');
}
