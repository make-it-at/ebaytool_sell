/**
 * eBay出品作業効率化ツール - ロガーモジュール
 * 
 * ログ記録機能を提供します。
 * 
 * バージョン: v1.3.1
 * 最終更新日: 2025-05-14
 */

// Logger名前空間
const Logger = {};

// 処理の開始時間を記録する変数
Logger.processStartTime = null;

/**
 * 通常ログを記録する
 * @param {string} message ログメッセージ
 */
Logger.log = function(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(Config.SHEET_NAMES.LOG);
  
  if (!logSheet) {
    console.log('ログシートが見つかりません: ' + message);
    return;
  }
  
  // タイムスタンプを取得
  const timestamp = new Date();
  
  // ログを追加
  logSheet.appendRow([
    Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
    'INFO',
    message
  ]);
};

/**
 * エラーログを記録する
 * @param {string} message エラーメッセージ
 */
Logger.logError = function(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(Config.SHEET_NAMES.LOG);
  
  if (!logSheet) {
    console.log('ログシートが見つかりません (ERROR): ' + message);
    return;
  }
  
  // タイムスタンプを取得
  const timestamp = new Date();
  
  // エラーログを追加
  logSheet.appendRow([
    Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
    'ERROR',
    message
  ]);
};

/**
 * 処理開始ログを記録する
 * @param {string} processName 処理名
 */
Logger.startProcess = function(processName) {
  // 処理開始時間を記録
  this.processStartTime = new Date();
  
  // ログを記録
  this.log('処理開始: ' + processName);
};

/**
 * 処理終了ログを記録する
 * @param {string} message 終了メッセージ
 */
Logger.endProcess = function(message) {
  // 処理時間を計算
  let processDuration = '不明';
  
  if (this.processStartTime) {
    const endTime = new Date();
    const durationMs = endTime.getTime() - this.processStartTime.getTime();
    const durationSec = Math.floor(durationMs / 1000);
    
    processDuration = durationSec + '秒';
    
    // 開始時間をリセット
    this.processStartTime = null;
  }
  
  // ログを記録
  this.log(message + ' (処理時間: ' + processDuration + ')');
};

/**
 * ログのクリアを行う
 * @param {number} daysToKeep 保持する日数
 */
Logger.clearOldLogs = function(daysToKeep = 30) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(Config.SHEET_NAMES.LOG);
  
  if (!logSheet || logSheet.getLastRow() <= 1) {
    return; // ログがないか、ヘッダーのみの場合
  }
  
  // 基準日を計算
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
  
  // ログデータを取得
  const dataRange = logSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行をスキップ
  const headerRow = values[0];
  const dataRows = values.slice(1);
  
  // 保持するログを選別
  const logsToKeep = [headerRow];
  
  dataRows.forEach(row => {
    const timestampStr = row[0];
    
    // タイムスタンプが日付形式でない場合はスキップ
    if (!timestampStr || typeof timestampStr === 'string' && !/\d{4}-\d{2}-\d{2}/.test(timestampStr)) {
      return;
    }
    
    const timestamp = new Date(timestampStr);
    
    // 基準日より新しいログのみ保持
    if (timestamp >= cutoffDate) {
      logsToKeep.push(row);
    }
  });
  
  // シートをクリアして保持するログを書き込み直す
  logSheet.clearContents();
  
  if (logsToKeep.length > 0) {
    logSheet.getRange(1, 1, logsToKeep.length, logsToKeep[0].length).setValues(logsToKeep);
  }
  
  this.log(`${daysToKeep}日以前のログを削除しました。${logsToKeep.length - 1}件のログを保持しています。`);
}; 