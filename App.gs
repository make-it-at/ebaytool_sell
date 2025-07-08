/**
 * eBayツール出品ファイル加工ツール - メインアプリケーション
 * 
 * メインアプリケーションの初期化、UI表示、イベントハンドリングを担当します。
 * 
 * バージョン: v1.5.13
 * 最終更新日: 2025-06-14
 * 更新内容: 処理完了メッセージから詳細情報を削除
 */

// アプリケーションのバージョン情報
const APP_VERSION = 'v1.5.13';

/**
 * eBayツール出品ファイル加工ツール
 * 
 * このスクリプトはeBay出品作業を効率化するためのGoogle Apps Scriptプロジェクトです。
 * 商品データの処理、フィルタリング、eBayフォーマットへの変換を自動化します。
 */

// グローバル関数としてUI関数をエクスポート（サイドバーからのアクセス用）
function showImportDialog() {
  UI.showImportDialog();
}

function showExportDialog() {
  UI.showExportDialog();
}

function showSettingsDialog() {
  UI.showSettingsDialog();
}

function showHelpDialog() {
  UI.showHelpDialog();
}

// グローバル関数としてFilters関数をエクスポート（サイドバーからのアクセス用）
function runNgWordFilter() {
  return Filters.runNgWordFilter();
}

function runDuplicateCheck() {
  return Filters.runDuplicateCheck();
}

function runLengthFilter() {
  return Filters.runLengthFilter();
}

function runLocationFix() {
  return Filters.runLocationFix();
}

function runPriceFilter() {
  return Filters.runPriceFilter();
}

// プログレスバーの状態をリセットする
function resetProgressState() {
  return UI.resetProgressState();
}

/**
 * HTMLファイルを含める
 * @param {string} filename 含めるHTMLファイル名
 * @return {string} HTMLコンテンツ
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 共通ユーティリティ関数

/**
 * 日付を「YYYY-MM-DD」形式でフォーマットする
 * @param {Date} date 日付
 * @return {string} フォーマットされた日付文字列
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * 日時を「YYYY-MM-DD HH:MM:SS」形式でフォーマットする
 * @param {Date} date 日時
 * @return {string} フォーマットされた日時文字列
 */
function formatDateTime(date) {
  const dateStr = formatDate(date);
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${dateStr} ${hours}:${minutes}:${seconds}`;
}

/**
 * 現在のタイムスタンプを取得する
 * @return {string} 現在の日時文字列
 */
function getCurrentTimestamp() {
  return formatDateTime(new Date());
}

/**
 * 文字列の前後の空白を除去する
 * @param {string} str 入力文字列
 * @return {string} トリムされた文字列
 */
function trimString(str) {
  if (!str) return '';
  return String(str).trim();
}

/**
 * 数値を通貨形式でフォーマットする
 * @param {number} value 数値
 * @param {string} currency 通貨コード（デフォルト: 'USD'）
 * @return {string} フォーマットされた通貨文字列
 */
function formatCurrency(value, currency = 'USD') {
  if (isNaN(value)) return '';
  
  try {
    return new Intl.NumberFormat('en-US', { 
      style: 'currency', 
      currency: currency,
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(value);
  } catch (e) {
    // フォールバックとして単純なフォーマットを使用
    return `$${value.toFixed(2)}`;
  }
}

// スプレッドシートが開かれたときに実行
function onOpen() {
  createMenu();
  showSidebar();
}

// カスタムメニューの作成
function createMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('eBayツール')
    .addItem('サイドバーを表示', 'showSidebar')
    .addSeparator()
    .addItem('ヘルプ', 'UI.showHelpDialog')
    .addToUi();
}

// サイドバーの表示
function showSidebar() {
  const template = HtmlService.createTemplateFromFile('Sidebar');
  
  // テンプレートにバージョン情報を渡す
  template.version = APP_VERSION;
  
  const html = template.evaluate()
    .setTitle('eBayツール出品ファイル加工ツール')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// すべての処理を順番に実行
function runAllProcesses() {
  // ログ記録開始
  Logger.startProcess('全処理一括実行');
  
  try {
    // 処理開始時に必ず前回の状態をリセットし、プログレスバーを表示
    resetProgressState();
    UI.showProgressBar('全処理一括実行を開始しています...', true);
    
    // プログレスバー更新を強制的に行うために短い遅延を入れる
    Utilities.sleep(300);
    
    // 各処理の配分（重要度や処理時間に応じて配分）
    const progressTracker = UI.ProgressTracker.init({
      preparation: 5,     // 準備段階
      ngWordFilter: 20,   // NGワードフィルター（重要度高）
      duplicateCheck: 15, // 重複チェック
      lengthFilter: 15,   // 文字数フィルター
      locationFix: 15,    // 所在地情報修正
      priceFilter: 20,    // 価格フィルター（重要度高）
      finalization: 10    // 最終処理
    });
    
    // 準備段階
    progressTracker.startStage('preparation', 'データの準備をしています...');
    
    // 処理前のデータ行数を取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // 出品データシートが存在するか確認
    if (!listingSheet) {
      throw new Error('出品データシートが見つかりません。初期設定を実行するか、データをインポートしてください。');
    }
    
    // 処理前のデータ行数を取得（ヘッダー行を除く）
    const beforeDataCount = listingSheet.getLastRow() - 1;
    
    // 各フィルター処理を順次実行する前に開始時間を保存
    const savedStartTime = Logger.processStartTime;
    
    // 結果を格納する変数を初期化
    const results = {
      ngWordFilter: { removed: 0, modified: 0 },
      duplicateCheck: { removed: 0 },
      lengthFilter: { removed: 0, limit: 0 },
      locationFix: { modified: 0 },
      priceFilter: { removed: 0, threshold: 0 }
    };
    
    // 準備段階完了
    progressTracker.completeStage();
    
    // NGワードフィルター実行
    progressTracker.startStage('ngWordFilter', 'NGワードフィルタリングを実行中...');
    const ngWordResult = Filters.runNgWordFilter();
    if (ngWordResult && ngWordResult.stats) {
      results.ngWordFilter.removed = ngWordResult.stats.removedCount || 0;
      results.ngWordFilter.modified = ngWordResult.stats.modifiedCount || 0;
    }
    progressTracker.completeStage();
    
    // 重複チェック実行
    progressTracker.startStage('duplicateCheck', '重複チェックを実行中...');
    const duplicateResult = Filters.runDuplicateCheck();
    if (duplicateResult && duplicateResult.stats) {
      results.duplicateCheck.removed = duplicateResult.stats.removedCount || 0;
    }
    progressTracker.completeStage();
    
    // 文字数フィルター実行
    progressTracker.startStage('lengthFilter', '文字数フィルタリングを実行中...');
    const lengthResult = Filters.runLengthFilter();
    if (lengthResult && lengthResult.stats) {
      results.lengthFilter.removed = lengthResult.stats.removedCount || 0;
      results.lengthFilter.limit = lengthResult.stats.characterLimit || 0;
    }
    progressTracker.completeStage();
    
    // 所在地情報修正実行
    progressTracker.startStage('locationFix', '所在地情報修正を実行中...');
    const locationResult = Filters.runLocationFix();
    if (locationResult && locationResult.stats) {
      results.locationFix.modified = locationResult.stats.modifiedCount || 0;
    }
    progressTracker.completeStage();
    
    // 価格フィルター実行
    progressTracker.startStage('priceFilter', '価格フィルタリングを実行中...');
    const priceResult = Filters.runPriceFilter();
    if (priceResult && priceResult.stats) {
      results.priceFilter.removed = priceResult.stats.removedCount || 0;
      results.priceFilter.threshold = priceResult.stats.priceThreshold || 0;
    }
    progressTracker.completeStage();
    
    // 最終処理
    progressTracker.startStage('finalization', '処理を完了しています...');
    
    // わずかな遅延を入れる（体感的な進行感のため）
    Utilities.sleep(200);
    
    // 開始時間を元に戻す
    Logger.processStartTime = savedStartTime;
    
    // 処理後のデータ行数を取得（ヘッダー行を除く）
    const afterDataCount = listingSheet.getLastRow() - 1;
    
    // 処理結果の詳細メッセージを作成
    const totalRemoved = results.ngWordFilter.removed + results.duplicateCheck.removed + 
                         results.lengthFilter.removed + results.priceFilter.removed;
    const totalModified = results.ngWordFilter.modified + results.locationFix.modified;
    
    // 詳細な追加情報を作成
    const additionalInfo = 
      `■ NGワード処理: ${results.ngWordFilter.removed}件削除, ${results.ngWordFilter.modified}件修正\n` +
      `■ 重複チェック: ${results.duplicateCheck.removed}件削除\n` +
      `■ 文字数制限(${results.lengthFilter.limit}文字以下): ${results.lengthFilter.removed}件削除\n` +
      `■ 所在地情報: ${results.locationFix.modified}件修正\n` +
      `■ 価格フィルター($${results.priceFilter.threshold}以下): ${results.priceFilter.removed}件削除`;
    
    // 完了表示
    progressTracker.complete('処理が完了しました');
    
    // 進捗を100%に更新（完了フラグはfalseのまま）
    UI.updateProgressBar(100);
    
    // 少し待機してからプログレスバーを非表示
    Utilities.sleep(1500);
    
    // プログレスバーを非表示
    UI.hideProgressBar();
    
    // 完了メッセージをUI.showResultMessageで表示
    UI.showResultMessage(
      'すべての処理が完了しました',
      {
        removedCount: totalRemoved,
        modifiedCount: totalModified,
        totalProcessed: beforeDataCount,
        beforeCount: beforeDataCount,
        afterCount: afterDataCount
      },
      additionalInfo
    );
    
    // 確実にメッセージを表示するため、LAST_RESULT_MESSAGEを直接設定
    LAST_RESULT_MESSAGE = {
      type: 'result',
      title: 'すべての処理が完了しました',
      stats: {
        removedCount: totalRemoved,
        modifiedCount: totalModified,
        totalProcessed: beforeDataCount,
        beforeCount: beforeDataCount,
        afterCount: afterDataCount
      },
      additionalInfo: additionalInfo
    };
    
    // ログ記録終了
    Logger.endProcess('全処理一括実行 成功');
    
    // 全処理一括実行完了フラグを設定
    setAllProcessesExecuted(true);
    
    // クライアント側表示用のメッセージを作成
    const clientSideMessage = `すべての処理が完了しました（データ数: ${beforeDataCount}件 → ${afterDataCount}件, 削除: ${totalRemoved}件, 修正: ${totalModified}件）\n\n${additionalInfo}`;
    
    return {
      success: true,
      message: clientSideMessage,
      stats: results
    };
  } catch (error) {
    // エラー発生時
    Logger.logError('全処理一括実行中にエラーが発生: ' + error.message);
    UI.showErrorMessage('処理中にエラーが発生しました: ' + error.message);
    
    return {
      success: false,
      message: '処理中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 全処理一括実行の実行状態を管理するフラグ
 * CSVインポート後にfalseにリセットされ、全処理一括実行完了後にtrueになる
 */
var ALL_PROCESSES_EXECUTED = false;

/**
 * 全処理一括実行の実行状態を取得
 * @return {boolean} 実行済みの場合true、未実行の場合false
 */
function isAllProcessesExecuted() {
  return ALL_PROCESSES_EXECUTED;
}

/**
 * 全処理一括実行の実行状態を設定
 * @param {boolean} executed 実行状態
 */
function setAllProcessesExecuted(executed) {
  ALL_PROCESSES_EXECUTED = executed;
  Logger.logInfo('全処理一括実行の実行状態を更新: ' + (executed ? '実行済み' : '未実行'));
}

/**
 * 全処理一括実行の実行状態をリセット（CSVインポート時に呼び出される）
 */
function resetAllProcessesExecuted() {
  ALL_PROCESSES_EXECUTED = false;
  Logger.logInfo('全処理一括実行の実行状態をリセット');
}

// スクリプト実行時の初期設定
function initialize() {
  // スプレッドシートの初期設定
  setupSheets();
}

// 必要なシートの設定
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = [
    Config.SHEET_NAMES.IMPORT, 
    Config.SHEET_NAMES.LISTING,  // 出品データシートを追加
    Config.SHEET_NAMES.SETTINGS, 
    Config.SHEET_NAMES.LOG,
    Config.SHEET_NAMES.TEST_DATA
  ];
  
  // 必要なシートが存在しない場合は作成
  sheetNames.forEach(name => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });
  
  // 各シートの初期設定
  setupImportSheet(ss.getSheetByName(Config.SHEET_NAMES.IMPORT));
  setupListingSheet(ss.getSheetByName(Config.SHEET_NAMES.LISTING));  // 出品データシートの設定
  setupSettingsSheet(ss.getSheetByName(Config.SHEET_NAMES.SETTINGS));
  setupLogSheet(ss.getSheetByName(Config.SHEET_NAMES.LOG));
  setupTestDataSheet(ss.getSheetByName(Config.SHEET_NAMES.TEST_DATA));
}

// データインポートシートの設定
function setupImportSheet(sheet) {
  // ヘッダー行の設定
  const headers = Config.SHEET_HEADERS.IMPORT;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

// 出品データシートの設定
function setupListingSheet(sheet) {
  // 出品データシートのヘッダー設定（テストデータと同じフォーマット）
  const headers = ["Action(CC=Cp1252)","CustomLabel","StartPrice","ConditionID","Title","Description","PicURL","UPC","Category","PaymentProfileName","ReturnProfileName","ShippingProfileName","Country","Location","Apply Profile Domestic","Apply Profile International","PayPalAccepted","PayPalEmailAddress","BuyerRequirements:LinkedPayPalAccount","Duration","Format","Quantity","Currency","SiteID","BestOfferEnabled"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

// 設定シートの設定
function setupSettingsSheet(sheet) {
  // 各セクションのヘッダー設定
  // NGワードセクション
  sheet.getRange("A1:B1").setValues([['NGワードリスト', '説明']]);
  sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#e6e6e6");
  
  // 初期NGワードの例
  const initialNgWords = [
    ['transformers', 'リスト全削除モードのテスト用NGワード'],
    ['transformer', ''],
    ['Transformers', ''],
    ['Transformer', ''],
    ['Django Unchain', ''],
    ['Django unchaine', ''],
    ['django unchaine', '']
  ];
  sheet.getRange(2, 1, initialNgWords.length, 2).setValues(initialNgWords);
  
  // 設定項目セクション (空行を入れて区切る)
  const ngWordsEndRow = 2 + initialNgWords.length;
  const settingsStartRow = ngWordsEndRow + 2; // 1行空けて開始
  
  sheet.getRange(settingsStartRow, 1, 1, 5).setValues([['設定項目', 'NGワードモード', '文字数制限', '価格下限', '重複類似度閾値']]);
  sheet.getRange(settingsStartRow, 1, 1, 5).setFontWeight("bold").setBackground("#e6e6e6");
  
  // 設定項目の初期値
  sheet.getRange(settingsStartRow + 1, 1, 1, 5).setValues([['値', 'リスト全削除', '20', '10', '80']]);
  sheet.getRange(settingsStartRow + 2, 1, 1, 5).setValues([['説明', 'NGワード処理方法。「リスト全削除」または「部分削除モード」', '商品名の最大文字数', '最低価格（ドル）', '重複判定の類似度閾値（%）']]);
  
  // 所在地置換パターンセクション (空行を入れて区切る)
  const locationStartRow = settingsStartRow + 4; // 設定項目の後、1行空けて開始
  
  sheet.getRange(locationStartRow, 1, 1, 3).setValues([['所在地置換パターン', '', '']]);
  sheet.getRange(locationStartRow, 1, 1, 3).setFontWeight("bold").setBackground("#e6e6e6");
  sheet.getRange(locationStartRow + 1, 1, 1, 3).setValues([['検索', '置換', '説明']]);
  sheet.getRange(locationStartRow + 1, 1, 1, 3).setFontWeight("bold").setBackground("#f2f2f2");
  
  // 初期パターンの設定
  const locationPatterns = [
    ['[0-9]+', '', '数字を削除します'],
    ['tokyo', '東京', '英語表記を日本語に変換'],
    ['osaka', '大阪', '英語表記を日本語に変換']
  ];
  
  sheet.getRange(locationStartRow + 2, 1, locationPatterns.length, 3).setValues(locationPatterns);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
}

// ログシートの設定
function setupLogSheet(sheet) {
  // ヘッダー行の設定
  const headers = ['タイムスタンプ', 'イベント', '詳細'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

// テストデータシートの設定
function setupTestDataSheet(sheet) {
  // ヘッダー行を設定 (必要に応じてeBay Export May 12 2025.csvのヘッダーに合わせる)
  const headers = ["Action(CC=Cp1252)","CustomLabel","StartPrice","ConditionID","Title","Description","PicURL","UPC","Category","PaymentProfileName","ReturnProfileName","ShippingProfileName","Country","Location","Apply Profile Domestic","Apply Profile International","PayPalAccepted","PayPalEmailAddress","BuyerRequirements:LinkedPayPalAccount","Duration","Format","Quantity","Currency","SiteID","BestOfferEnabled","",""];
  
  // シートがすでにデータを持っているか確認
  const existingData = sheet.getDataRange().getValues();
  
  // すでにデータがあり、ヘッダーも一致する場合は何もしない
  if (existingData.length > 1 && arraysEqual(existingData[0], headers)) {
    return;
  }
  
  // シートをクリアして新しいヘッダーを設定
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  
  // テストデータのパスとファイル名の情報を表示
  sheet.getRange(2, 1, 1, 3).setValues([["テストデータ情報", "パス:", "/Users/kyosukemakita/Documents/Cursor/ebaytool_sell/eBay Export May 12 2025.csv"]]);
  sheet.getRange(3, 1, 1, 3).setValues([["注意事項", "", "このシートにはテスト用CSVデータが入ります。データを「データインポート」シートに転記するには「テスト用CSVインポート」ボタンを使用してください。"]]);
  
  // 見やすくするために書式設定
  sheet.getRange(2, 1, 2, 3).setBackground("#f2f2f2");
  sheet.getRange(2, 1, 2, 1).setFontWeight("bold");
  
  // サンプルデータがないことを通知
  sheet.getRange(5, 1).setValue("テストデータがまだインポートされていません。CSVファイルを手動でこのシートに貼り付けるか、設定してください。");
  sheet.getRange(5, 1).setFontColor("red");
}

/**
 * 配列が等しいかどうかをチェックするヘルパー関数
 */
function arraysEqual(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

/**
 * Base64エンコードされたCSVデータをインポートする
 * クライアント側JSからのエントリーポイント
 * @param {string} base64Data Base64エンコードされたCSVデータ
 * @return {boolean} インポート成功フラグ
 */
function importCsvFromBase64(base64Data) {
  try {
    // Base64デコード
    const decodedData = Utilities.base64Decode(base64Data);
    
    // UTF-8のバイト配列を文字列に変換
    const csvData = Utilities.newBlob(decodedData).getDataAsString();
    
    // CSVデータをBlobとして作成
    const blob = Utilities.newBlob(csvData, 'text/csv', 'import.csv');
    
    // ImportExportモジュールの関数を呼び出し
    const result = ImportExport.importCsv(blob);
    if (result) {
      // サイドバーに処理結果を表示
      UI.showResultMessage(
        'CSVインポート完了',
        { removedCount: 0, modifiedCount: 0, totalProcessed: 0 },
        'CSVファイルのインポートが正常に完了しました。データが「出品データ」シートに反映されています。'
      );
      // 完了時に進捗を100%に
      UI.updateProgressBar(100);
    }
    return result;
  } catch (error) {
    Logger.logError('CSVインポート(Base64)でエラーが発生: ' + error.message);
    return false;
  }
}

/**
 * CSVとしてエクスポートする
 * クライアント側JSからのエントリーポイント
 * @param {string} sheetName エクスポート対象のシート名（オプション）
 * @return {string} ダウンロード用URL
 */
function exportToCsv(sheetName) {
  try {
    return ImportExport.exportToCsv(sheetName);
  } catch (error) {
    Logger.logError('CSVエクスポート中にエラーが発生: ' + error.message);
    return null;
  }
}

/**
 * サイドバーからの直接CSVエクスポート
 * クライアント側JSからのエントリーポイント
 * @param {string} sheetName エクスポート対象のシート名（オプション）
 * @return {Object} エクスポートの結果（成功: true, csvData: CSVデータ, fileName: ファイル名）
 */
function exportCsvFromSidebar(sheetName) {
  try {
    const result = ImportExport.exportToCsv(sheetName);
    
    // シートからデータ行数を取得
    let rowCount = 0;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName || Config.SHEET_NAMES.LISTING);
    
    if (sheet) {
      // ヘッダー行を除いた行数
      rowCount = Math.max(0, sheet.getLastRow() - 1);
    }
    
    if (result) {
      return {
        success: true,
        csvData: result.csvData,
        fileName: result.fileName,
        rowCount: rowCount  // データ行数を追加
      };
    } else {
      return {
        success: false,
        message: 'エクスポートに失敗しました。'
      };
    }
  } catch (error) {
    Logger.logError('サイドバーからのCSVエクスポート中にエラーが発生: ' + error.message);
    return {
      success: false,
      message: error.message || 'エクスポート中にエラーが発生しました。'
    };
  }
}

/**
 * NGワード設定シートを開く
 * クライアント側JSからのエントリーポイント
 */
function openNgWordsSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(Config.SHEET_NAMES.SETTINGS);
    
    if (!settingsSheet) {
      UI.showErrorMessage('設定シートが見つかりません。初期化を実行してください。');
      return false;
    }
    
    // 設定シートをアクティブにする
    ss.setActiveSheet(settingsSheet);
    
    // NGワードの位置にカーソルを移動
    settingsSheet.getRange(1, 1).activate();
    
    // 成功メッセージ
    UI.showSuccessMessage('NGワード設定を開きました。設定シートのNGワードリストを編集してください。');
    
    return true;
  } catch (error) {
    Logger.logError('NGワード設定を開く際にエラーが発生: ' + error.message);
    UI.showErrorMessage('NGワード設定を開けませんでした: ' + error.message);
    return false;
  }
}

/**
 * テストデータシートから出品データシートにデータを転記する
 * テスト用CSVインポート機能
 */
function importTestCsv() {
  Logger.startProcess('テスト用CSVインポート');
  UI.showProgressBar('テスト用CSVデータをインポートしています...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const testDataSheet = ss.getSheetByName(Config.SHEET_NAMES.TEST_DATA);
    let listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // テストデータシートが存在するか確認
    if (!testDataSheet) {
      throw new Error('テストデータシートが見つかりません。初期設定を実行してください。');
    }
    
    // 出品データシートが存在しない場合は作成
    if (!listingSheet) {
      listingSheet = ss.insertSheet(Config.SHEET_NAMES.LISTING);
      setupListingSheet(listingSheet);
    }
    
    // テストデータシートからデータを取得
    const testData = testDataSheet.getDataRange().getValues();
    
    // データが十分にあるか確認
    if (testData.length <= 5) { // ヘッダー + 説明行を考慮
      throw new Error('テストデータシートに有効なデータがありません。CSVデータをシートに貼り付けてください。\n(パス: /Users/kyosukemakita/Documents/Cursor/ebaytool_sell/eBay Export May 12 2025.csv)');
    }
    
    // 実際のデータ部分を取得（ヘッダー行と説明行をスキップ）
    const headerRow = testData[0]; // ヘッダー行
    const actualData = testData.slice(5); // 説明行をスキップして実データを取得
    
    // データ行数チェック
    if (actualData.length > Config.MAX_ROWS) {
      throw new Error(`インポートするデータが多すぎます。最大${Config.MAX_ROWS}行までです。`);
    }
    
    // 出品データシートをクリアしてヘッダー行を設定
    listingSheet.clearContents();
    
    UI.updateProgressBar(30);
    
    // ヘッダー行を設定
    listingSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    
    UI.updateProgressBar(50);
    
    // 実データ部分を転記
    if (actualData.length > 0) {
      listingSheet.getRange(2, 1, actualData.length, headerRow.length).setValues(actualData);
    }
    
    UI.updateProgressBar(90);
    
    // 完了メッセージの表示
    const message = `テスト用CSVデータが正常にインポートされました。\n${actualData.length}件のデータを出品データシートにインポートしました。\nファイル: eBay Export May 12 2025.csv`;
    UI.showSuccessMessage(message);
    Logger.endProcess('テスト用CSVインポート完了');
    
    UI.updateProgressBar(100);
    
    return true;
  } catch (error) {
    Logger.logError('テスト用CSVインポート中にエラー: ' + error.message);
    UI.showErrorMessage('テスト用CSVインポート中にエラーが発生しました: ' + error.message);
    return false;
  } finally {
    UI.hideProgressBar();
  }
}

// 以下の関数はサイドバーからの直接編集が不要になったためコメントアウト
// /**
//  * NGワードのリストを取得する
//  * クライアント側JSからのエントリーポイント
//  * @return {Array} NGワードのリスト
//  */
// function getNgWordsList() {
//   try {
//     const settings = Config.getSettings();
//     return settings.ngWords || [];
//   } catch (error) {
//     Logger.logError('NGワードリスト取得中にエラーが発生: ' + error.message);
//     return [];
//   }
// }

// /**
//  * NGワードのリストを保存する
//  * クライアント側JSからのエントリーポイント
//  * @param {Array} ngWords NGワードの配列
//  * @return {boolean} 保存成功フラグ
//  */
// function saveNgWordsList(ngWords) {
//   try {
//     return Config.saveSettings({
//       ngWords: ngWords
//     });
//   } catch (error) {
//     Logger.logError('NGワードリスト保存中にエラーが発生: ' + error.message);
//     return false;
//   }
// } 