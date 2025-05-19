/**
 * eBayツール出品ファイル加工ツール - メインアプリケーション
 * 
 * メインアプリケーションの初期化、UI表示、イベントハンドリングを担当します。
 * 
 * バージョン: v1.4.10
 * 最終更新日: 2025-05-28
 */

// アプリケーションのバージョン情報
const APP_VERSION = 'v1.4.10';

/**
 * eBayツール出品ファイル加工ツール
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

// すべての処理を順番に実行
function runAllProcesses() {
  // ログ記録開始
  Logger.startProcess('全処理一括実行');
  
  try {
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
    
    // NGワードフィルター実行
    const ngWordResult = Filters.runNgWordFilter();
    if (ngWordResult && ngWordResult.stats) {
      results.ngWordFilter.removed = ngWordResult.stats.removedCount || 0;
      results.ngWordFilter.modified = ngWordResult.stats.modifiedCount || 0;
    }
    
    // 重複チェック実行
    const duplicateResult = Filters.runDuplicateCheck();
    if (duplicateResult && duplicateResult.stats) {
      results.duplicateCheck.removed = duplicateResult.stats.removedCount || 0;
    }
    
    // 文字数フィルター実行
    const lengthResult = Filters.runLengthFilter();
    if (lengthResult && lengthResult.stats) {
      results.lengthFilter.removed = lengthResult.stats.removedCount || 0;
      results.lengthFilter.limit = lengthResult.stats.characterLimit || 0;
    }
    
    // 所在地情報修正実行
    const locationResult = Filters.runLocationFix();
    if (locationResult && locationResult.stats) {
      results.locationFix.modified = locationResult.stats.modifiedCount || 0;
    }
    
    // 価格フィルター実行
    const priceResult = Filters.runPriceFilter();
    if (priceResult && priceResult.stats) {
      results.priceFilter.removed = priceResult.stats.removedCount || 0;
      results.priceFilter.threshold = priceResult.stats.priceThreshold || 0;
    }
    
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
    
    // ログ記録終了
    Logger.endProcess('全処理一括実行 成功');
    
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
    return ImportExport.importCsv(blob);
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
    if (result) {
      return {
        success: true,
        csvData: result.csvData,
        fileName: result.fileName
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