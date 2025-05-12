/**
 * eBay出品作業効率化ツール
 * 
 * このスクリプトはeBay出品作業を効率化するためのGoogle Apps Scriptプロジェクトです。
 * 商品データの処理、フィルタリング、eBayフォーマットへの変換を自動化します。
 */

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
    .addItem('CSVインポート', 'UI.showImportDialog')
    .addItem('CSVエクスポート', 'UI.showExportDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('フィルター処理')
      .addItem('NGワードフィルタリング', 'Filters.runNgWordFilter')
      .addItem('重複チェック', 'Filters.runDuplicateCheck')
      .addItem('文字数制限フィルター', 'Filters.runLengthFilter')
      .addItem('所在地情報修正', 'Filters.runLocationFix')
      .addItem('価格フィルタリング', 'Filters.runPriceFilter')
    )
    .addSeparator()
    .addItem('全処理を一括実行', 'runAllProcesses')
    .addSeparator()
    .addItem('設定', 'UI.showSettingsDialog')
    .addItem('ヘルプ', 'UI.showHelpDialog')
    .addToUi();
}

// サイドバーの表示
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('eBay出品ツール')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

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

// すべての処理を順番に実行
function runAllProcesses() {
  // ログ記録開始
  Logger.startProcess('全処理一括実行');
  
  try {
    // 各フィルター処理を順次実行
    Filters.runNgWordFilter();
    Filters.runDuplicateCheck();
    Filters.runLengthFilter();
    Filters.runLocationFix();
    Filters.runPriceFilter();
    
    // 完了メッセージを表示
    UI.showSuccessMessage('すべての処理が完了しました。');
    
    // ログ記録終了
    Logger.endProcess('全処理一括実行 成功');
  } catch (error) {
    // エラー発生時
    Logger.logError('全処理一括実行中にエラーが発生: ' + error.message);
    UI.showErrorMessage('処理中にエラーが発生しました: ' + error.message);
  }
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
    Config.SHEET_NAMES.SETTINGS, 
    Config.SHEET_NAMES.LOG
  ];
  
  // 必要なシートが存在しない場合は作成
  sheetNames.forEach(name => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });
  
  // 各シートの初期設定
  setupImportSheet(ss.getSheetByName(Config.SHEET_NAMES.IMPORT));
  setupSettingsSheet(ss.getSheetByName(Config.SHEET_NAMES.SETTINGS));
  setupLogSheet(ss.getSheetByName(Config.SHEET_NAMES.LOG));
}

// データインポートシートの設定
function setupImportSheet(sheet) {
  // ヘッダー行の設定
  const headers = Config.SHEET_HEADERS.IMPORT;
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
  sheet.getRange(settingsStartRow + 1, 1, 1, 5).setValues([['値', 'リスト全削除', '80', '10', '80']]);
  sheet.getRange(settingsStartRow + 2, 1, 1, 5).setValues([['説明', 'NGワード処理方法。「リスト全削除」または「部分削除モード」', '商品名の文字数上限', '最低価格（ドル）', '重複判定の類似度閾値（%）']]);
  
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
    return ImportExport.importCsv(blob);
  } catch (error) {
    Logger.logError('CSVインポート(Base64)でエラーが発生: ' + error.message);
    return false;
  }
}

/**
 * CSVとしてエクスポートする
 * クライアント側JSからのエントリーポイント
 * @return {string} ダウンロード用URL
 */
function exportToCsv() {
  try {
    return ImportExport.exportToCsv();
  } catch (error) {
    Logger.logError('CSVエクスポート中にエラーが発生: ' + error.message);
    return null;
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