/**
 * BigQuery連携 - スプレッドシートからBigQueryへの定期バッチ転送
 *
 * 【セットアップ手順】
 * 1. GASエディタで「サービス」→「BigQuery API」を追加
 * 2. 下記の定数を実際の値に変更
 * 3. GASエディタで「トリガー」→ createSyncTrigger() を1回実行してタイマー設定
 *    または手動で「トリガーを追加」→ syncToBigQuery → 時間主導型 → 5分おき
 *
 * 【BigQuery テーブル作成SQL（初回のみBQコンソールで実行）】
 *
 * CREATE TABLE `your-project.kanban_app.verification_records` (
 *   timestamp TIMESTAMP,
 *   employee_code STRING,
 *   shipping_date DATE,
 *   shipping_bin STRING,
 *   sebango_no STRING,
 *   product_model STRING,
 *   kanban_edaban STRING,
 *   setsudan_no_1 STRING,
 *   setsudan_no_2 STRING,
 *   setsudan_no_3 STRING,
 *   synced_at TIMESTAMP
 * );
 */

// ==== 設定（実際の値に変更してください） ====
var BQ_PROJECT_ID = 'YOUR_GCP_PROJECT_ID';
var BQ_DATASET_ID = 'kanban_app';
var BQ_TABLE_ID = 'verification_records';

// スプレッドシートの転送済みフラグ列（K列 = 11列目）
var SYNC_FLAG_COL = 11;

/**
 * スプレッドシートの未転送データをBigQueryに挿入する
 * トリガーで5分おきに実行
 */
function syncToBigQuery() {
  var ss = getDataSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_DATA);
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) return; // ヘッダーのみ

  var rowsToInsert = [];
  var rowIndices = [];

  for (var i = 1; i < data.length; i++) {
    var syncFlag = data[i][SYNC_FLAG_COL - 1];
    if (syncFlag === 'synced') continue; // 転送済みはスキップ

    var row = data[i];
    var timestamp = row[0] ? new Date(row[0]) : new Date();

    rowsToInsert.push({
      json: {
        timestamp: timestamp.toISOString(),
        employee_code: String(row[1] || ''),
        shipping_date: String(row[2] || ''),
        shipping_bin: String(row[3] || ''),
        sebango_no: String(row[4] || ''),
        product_model: String(row[5] || ''),
        kanban_edaban: String(row[6] || ''),
        setsudan_no_1: String(row[7] || ''),
        setsudan_no_2: String(row[8] || ''),
        setsudan_no_3: String(row[9] || ''),
        synced_at: new Date().toISOString()
      }
    });
    rowIndices.push(i);
  }

  if (rowsToInsert.length === 0) {
    Logger.log('転送するデータがありません');
    return;
  }

  // BigQuery Streaming Insert
  try {
    var insertAllRequest = { rows: rowsToInsert };

    var response = BigQuery.Tabledata.insertAll(
      insertAllRequest,
      BQ_PROJECT_ID,
      BQ_DATASET_ID,
      BQ_TABLE_ID
    );

    // エラーチェック
    if (response.insertErrors && response.insertErrors.length > 0) {
      Logger.log('BigQuery挿入エラー: ' + JSON.stringify(response.insertErrors));
      return;
    }

    // 転送成功 → フラグ更新
    for (var j = 0; j < rowIndices.length; j++) {
      sheet.getRange(rowIndices[j] + 1, SYNC_FLAG_COL).setValue('synced');
    }

    Logger.log(rowsToInsert.length + '件をBigQueryに転送しました');
  } catch (e) {
    Logger.log('BigQuery転送エラー: ' + e.message);
  }
}

/**
 * 5分おきの同期トリガーを作成する（初回のみ実行）
 */
function createSyncTrigger() {
  // 既存のトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'syncToBigQuery') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 5分おきのトリガーを作成
  ScriptApp.newTrigger('syncToBigQuery')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('BigQuery同期トリガーを作成しました（5分間隔）');
}

/**
 * 手動で全データを再転送する（リカバリ用）
 */
function resyncAllToBigQuery() {
  var ss = getDataSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_DATA);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return;

  // 全行のsyncフラグをクリア
  var range = sheet.getRange(2, SYNC_FLAG_COL, lastRow - 1, 1);
  range.clearContent();

  // 通常の同期を実行
  syncToBigQuery();
}
