/**
 * eBayツール出品ファイル加工ツール - メインエントリーポイント
 * 
 * GASアプリケーションのメインエントリーポイントとして機能し、
 * 各モジュールの初期化とグローバル関数のエクスポートを行います。
 * 
 * バージョン: v1.5.4
 * 最終更新日: 2025-06-16
 * 更新内容: 全処理一括実行の処理順序とエラーハンドリングを改善
 */

// アプリケーションの起動時に実行される
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
    .addItem('CSVインポート', 'showImportDialog')
    .addItem('CSVエクスポート', 'showExportDialog')
    .addSeparator()
    .addItem('設定', 'showSettingsDialog')
    .addItem('ヘルプ', 'showHelpDialog')
    .addToUi();
}

// サイドバーの表示
function showSidebar() {
  UI.showSidebar();
}

// モジュール関数へのグローバルアクセスを提供

// UI関連のグローバル関数
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

// フィルタリング関連のグローバル関数
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

// 全処理一括実行
function runAllProcesses() {
  try {
    // ログ記録開始
    Logger.startProcess('全処理一括実行');
    
    // 処理開始時に必ず前回の状態をリセットし、プログレスバーを表示
    resetProgressState();
    
    // 安定した表示のためにプログレスバーの状態を明示的に設定
    _progressState = {
      isVisible: true,
      message: '全処理一括実行を開始しています...',
      percent: 0,
      completion: false
    };
    
    // 処理開始のメッセージを表示
    UI.showProgressBar('全処理一括実行を開始しています...', true);
    
    // プログレスバー更新を強制的に行うために短い遅延を入れる
    Utilities.sleep(500);
    
    // 各処理の配分（重要度や処理時間に応じて配分）
    const progressTracker = UI.ProgressTracker.init({
      preparation: 10,    // 準備段階
      ngWordFilter: 20,   // NGワードフィルター（重要度高）
      duplicateCheck: 15, // 重複チェック
      lengthFilter: 15,   // 文字数フィルター
      locationFix: 15,    // 所在地情報修正
      priceFilter: 15,    // 価格フィルター（重要度高）
      finalization: 10    // 最終処理
    });
    
    // 準備段階
    progressTracker.startStage('preparation', 'データの準備をしています...');
    
    // スプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // 出品データシートが存在するか確認
    if (!listingSheet) {
      // シートが存在しない場合は作成
      listingSheet = ss.insertSheet(Config.SHEET_NAMES.LISTING);
      
      // ヘッダー行を設定
      const defaultHeaders = ["Action(CC=Cp1252)","CustomLabel","StartPrice","ConditionID","Title","Description","PicURL","UPC","Category","PaymentProfileName","ReturnProfileName","ShippingProfileName","Country","Location","Apply Profile Domestic","Apply Profile International","PayPalAccepted","PayPalEmailAddress","BuyerRequirements:LinkedPayPalAccount","Duration","Format","Quantity","Currency","SiteID","BestOfferEnabled"];
      listingSheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
      listingSheet.setFrozenRows(1);
      
      Logger.logInfo('出品データシートが見つからなかったため、新規作成しました。');
    }
    
    // データ行の確認
    const lastRow = listingSheet.getLastRow();
    const dataExists = lastRow > 1; // ヘッダー行以外のデータがあるか
    
    // 処理前のデータ行数を取得（ヘッダー行を除く）
    const beforeDataCount = Math.max(0, lastRow - 1);
    
    // データがない場合は警告
    if (!dataExists) {
      Logger.logWarning('データが存在しません。先にCSVをインポートしてください。');
      UI.showWarningMessage('データが存在しません。先にCSVをインポートしてください。全処理一括実行を中止します。');
      
      // 処理を終了
      progressTracker.complete('データがないため処理を終了します');
      UI.hideProgressBar();
      return {
        success: false,
        message: 'データが存在しないため処理を中止しました。先にCSVをインポートしてください。'
      };
    }
    
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
    
    // ヘッダー行の取得
    const headerRow = listingSheet.getRange(1, 1, 1, listingSheet.getLastColumn()).getValues()[0];
    
    // 必要なカラムが存在するか確認
    const titleColumnExists = headerRow.indexOf('Title') !== -1;
    const locationColumnExists = headerRow.indexOf('Location') !== -1 || headerRow.indexOf('所在地') !== -1;
    const priceColumnExists = headerRow.indexOf('StartPrice') !== -1;
    
    // 準備段階完了
    progressTracker.completeStage();
    
    // 1. NGワードフィルター実行
    progressTracker.startStage('ngWordFilter', 'NGワードフィルタリングを実行中...');
    
    if (titleColumnExists) {
    const ngWordResult = Filters.runNgWordFilter();
    if (ngWordResult && ngWordResult.stats) {
      results.ngWordFilter.removed = ngWordResult.stats.removedCount || 0;
      results.ngWordFilter.modified = ngWordResult.stats.modifiedCount || 0;
    }
    } else {
      Logger.logWarning('Title列が見つからないため、NGワードフィルターをスキップします');
    }
    
    progressTracker.completeStage();
    
    // 2. 重複チェック実行
    progressTracker.startStage('duplicateCheck', '重複チェックを実行中...');
    
    if (titleColumnExists) {
    const duplicateResult = Filters.runDuplicateCheck();
    if (duplicateResult && duplicateResult.stats) {
      results.duplicateCheck.removed = duplicateResult.stats.removedCount || 0;
    }
    } else {
      Logger.logWarning('Title列が見つからないため、重複チェックをスキップします');
    }
    
    progressTracker.completeStage();
    
    // 3. 文字数フィルター実行
    progressTracker.startStage('lengthFilter', '文字数フィルタリングを実行中...');
    
    if (titleColumnExists) {
    const lengthResult = Filters.runLengthFilter();
    if (lengthResult && lengthResult.stats) {
      results.lengthFilter.removed = lengthResult.stats.removedCount || 0;
      results.lengthFilter.limit = lengthResult.stats.characterLimit || 0;
    }
    } else {
      Logger.logWarning('Title列が見つからないため、文字数フィルターをスキップします');
    }
    
    progressTracker.completeStage();
    
    // 4. 所在地情報修正実行
    progressTracker.startStage('locationFix', '所在地情報修正を実行中...');
    
    if (locationColumnExists) {
    const locationResult = Filters.runLocationFix();
    if (locationResult && locationResult.stats) {
      results.locationFix.modified = locationResult.stats.modifiedCount || 0;
    }
    } else {
      Logger.logWarning('Location列または所在地列が見つからないため、所在地情報修正をスキップします');
    }
    
    progressTracker.completeStage();
    
    // 5. 価格フィルター実行
    progressTracker.startStage('priceFilter', '価格フィルタリングを実行中...');
    
    if (priceColumnExists) {
    const priceResult = Filters.runPriceFilter();
    if (priceResult && priceResult.stats) {
      results.priceFilter.removed = priceResult.stats.removedCount || 0;
      results.priceFilter.threshold = priceResult.stats.priceThreshold || 0;
    }
    } else {
      Logger.logWarning('StartPrice列が見つからないため、価格フィルターをスキップします');
    }
    
    progressTracker.completeStage();
    
    // 最終処理
    progressTracker.startStage('finalization', '処理を完了しています...');
    
    // わずかな遅延を入れる（体感的な進行感のため）
    Utilities.sleep(200);
    
    // 開始時間を元に戻す
    Logger.processStartTime = savedStartTime;
    
    // 処理後のデータ行数を取得（ヘッダー行を除く）
    const afterDataCount = Math.max(0, listingSheet.getLastRow() - 1);
    
    // 処理結果の詳細メッセージを作成
    const totalRemoved = results.ngWordFilter.removed + results.duplicateCheck.removed + 
                        results.lengthFilter.removed + results.priceFilter.removed;
    const totalModified = results.ngWordFilter.modified + results.locationFix.modified;
    
    // スキップされた処理があるかどうか
    const hasSkippedProcess = !titleColumnExists || !locationColumnExists || !priceColumnExists;
    
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
      }
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
      }
    };
    
    // スキップされた処理がある場合は警告メッセージを追加
    if (hasSkippedProcess) {
      let warningMessage = '警告: 一部の処理がスキップされました。';
      if (!titleColumnExists) warningMessage += ' Title列が見つかりません。';
      if (!locationColumnExists) warningMessage += ' Location/所在地列が見つかりません。';
      if (!priceColumnExists) warningMessage += ' StartPrice列が見つかりません。';
      
      UI.showWarningMessage(warningMessage);
      Logger.logWarning(warningMessage);
    }
    
    // ログ記録終了
    Logger.endProcess('全処理一括実行 成功');
    
    // クライアント側表示用のメッセージを作成
    let clientSideMessage = `すべての処理が完了しました（データ数: ${beforeDataCount}件 → ${afterDataCount}件, 削除: ${totalRemoved}件）`;
    
    return {
      success: true,
      message: clientSideMessage,
      stats: results,
      hasSkippedProcess: hasSkippedProcess
    };
  } catch (error) {
    // エラー発生時
    Logger.logError('全処理一括実行中にエラーが発生: ' + error.message);
    UI.showErrorMessage('処理中にエラーが発生しました: ' + error.message);
    
    // プログレスバーを非表示
    UI.hideProgressBar();
    
    return {
      success: false,
      message: '処理中にエラーが発生しました: ' + error.message
    };
  }
}

// CSVインポート処理のグローバル関数
function importCsv(csvFile) {
  return ImportExport.importCsv(csvFile);
}

// CSVエクスポート処理のグローバル関数
function exportToCsv() {
  return ImportExport.exportToCsv();
}

// eBay形式のCSVエクスポート処理のグローバル関数
function exportToEbayFormat() {
  return ImportExport.formatForEbay();
}

// プログレスバーの状態をリセットする
function resetProgressState() {
  return UI.resetProgressState();
}

// 設定関連のグローバル関数
function saveSettings(settings) {
  return Config.saveSettings(settings);
}

function getSettings() {
  return Config.getSettings();
}

// 初期設定を実行
function initializeApp() {
  return Config.initializeApp();
}

// テストデータをインポート（開発・テスト用）
function importTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testDataSheet = ss.getSheetByName(Config.SHEET_NAMES.TEST_DATA);
  
  if (!testDataSheet) {
    UI.showErrorMessage('テストデータシートが見つかりません。テストデータシートを作成し、データを設定してください。');
    return false;
  }
  
  // テストデータの範囲を取得（6行目以降がデータ、1-5行目は情報）
  const dataStartRow = 6;
  const lastRow = testDataSheet.getLastRow();
  const lastCol = testDataSheet.getLastColumn();
  
  if (lastRow < dataStartRow) {
    UI.showErrorMessage('テストデータが見つかりません。テストデータシートにデータを設定してください。');
    return false;
  }
  
  // データを取得
  const testData = testDataSheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol).getValues();
  
  // インポート処理を実行
  Logger.startProcess('テストデータインポート');
  UI.showProgressBar('テストデータをインポートしています...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // シートが存在しない場合は作成
    if (!listingSheet) {
      listingSheet = ss.insertSheet(Config.SHEET_NAMES.LISTING);
      Logger.logInfo('出品データシートが見つからなかったため、新規作成しました。');
    }
    
    // シートをクリア
    if (listingSheet.getLastRow() > 0) {
      listingSheet.clearContents();
    }
    
    // ヘッダー行を設定
    const headers = testData[0];
    listingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    listingSheet.setFrozenRows(1);
    
    // データ行を設定（ヘッダー行をスキップ）
    const dataRows = testData.slice(1);
    if (dataRows.length > 0) {
      listingSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    }
    
    UI.showSuccessMessage(`テストデータのインポートが完了しました。${dataRows.length}件のデータをインポートしました。`);
    Logger.endProcess('テストデータインポート完了');
    
    return true;
  } catch (error) {
    Logger.logError('テストデータインポート中にエラーが発生: ' + error.message);
    UI.showErrorMessage('テストデータインポート中にエラーが発生しました: ' + error.message);
    return false;
  } finally {
    UI.hideProgressBar();
  }
}

// テストデータ設定ヘルプを表示
function showTestDataHelp() {
  UI.showTestDataHelpDialog();
}

/**
 * データを初期化する（出品データとログをクリア）
 * サイドバーからの呼び出し用
 * @return {Object} 初期化結果
 */
function initializeData() {
  try {
    Logger.startProcess('データ初期化');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let isSuccess = true;
    let message = '';
    
    // 出品データシートを取得
    const listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // ログシートを取得
    const logSheet = ss.getSheetByName(Config.SHEET_NAMES.LOG);
    
    // 出品データシートの初期化
    if (listingSheet) {
      try {
        // ヘッダー行を保持
        const headers = listingSheet.getRange(1, 1, 1, listingSheet.getLastColumn()).getValues();
        
        // シートをクリア
        listingSheet.clear();
        
        // ヘッダー行を再設定
        if (headers && headers[0] && headers[0].length > 0) {
          listingSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
          listingSheet.setFrozenRows(1);
        } else {
          // ヘッダーがない場合はデフォルトヘッダーを設定
          const defaultHeaders = ["Action(CC=Cp1252)","CustomLabel","StartPrice","ConditionID","Title","Description","PicURL","UPC","Category","PaymentProfileName","ReturnProfileName","ShippingProfileName","Country","Location","Apply Profile Domestic","Apply Profile International","PayPalAccepted","PayPalEmailAddress","BuyerRequirements:LinkedPayPalAccount","Duration","Format","Quantity","Currency","SiteID","BestOfferEnabled"];
          listingSheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
          listingSheet.setFrozenRows(1);
        }
        
        Logger.log('出品データシートを初期化しました');
        message += '出品データシートを初期化しました。\n';
      } catch (e) {
        isSuccess = false;
        Logger.logError('出品データシート初期化エラー: ' + e.message);
        message += '出品データシート初期化中にエラーが発生しました: ' + e.message + '\n';
      }
    } else {
      // 出品データシートが存在しない場合は作成
      try {
        const newSheet = ss.insertSheet(Config.SHEET_NAMES.LISTING);
        const defaultHeaders = ["Action(CC=Cp1252)","CustomLabel","StartPrice","ConditionID","Title","Description","PicURL","UPC","Category","PaymentProfileName","ReturnProfileName","ShippingProfileName","Country","Location","Apply Profile Domestic","Apply Profile International","PayPalAccepted","PayPalEmailAddress","BuyerRequirements:LinkedPayPalAccount","Duration","Format","Quantity","Currency","SiteID","BestOfferEnabled"];
        newSheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
        newSheet.setFrozenRows(1);
        
        Logger.log('出品データシートを新規作成しました');
        message += '出品データシートを新規作成しました。\n';
      } catch (e) {
        isSuccess = false;
        Logger.logError('出品データシート作成エラー: ' + e.message);
        message += '出品データシート作成中にエラーが発生しました: ' + e.message + '\n';
      }
    }
    
    // ログシートの初期化
    if (logSheet) {
      try {
        // ヘッダー行を保持
        const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues();
        
        // シートをクリア
        logSheet.clear();
        
        // ヘッダー行を再設定
        if (logHeaders && logHeaders[0] && logHeaders[0].length > 0) {
          logSheet.getRange(1, 1, 1, logHeaders[0].length).setValues(logHeaders);
          logSheet.setFrozenRows(1);
        } else {
          // ヘッダーがない場合はデフォルトヘッダーを設定
          const defaultLogHeaders = ['タイムスタンプ', 'イベント', '詳細'];
          logSheet.getRange(1, 1, 1, defaultLogHeaders.length).setValues([defaultLogHeaders]);
          logSheet.setFrozenRows(1);
        }
        
        Logger.log('ログシートを初期化しました');
        message += 'ログシートを初期化しました。\n';
      } catch (e) {
        isSuccess = false;
        Logger.logError('ログシート初期化エラー: ' + e.message);
        message += 'ログシート初期化中にエラーが発生しました: ' + e.message + '\n';
      }
    } else {
      // ログシートが存在しない場合は作成
      try {
        const newLogSheet = ss.insertSheet(Config.SHEET_NAMES.LOG);
        const defaultLogHeaders = ['タイムスタンプ', 'イベント', '詳細'];
        newLogSheet.getRange(1, 1, 1, defaultLogHeaders.length).setValues([defaultLogHeaders]);
        newLogSheet.setFrozenRows(1);
        
        Logger.log('ログシートを新規作成しました');
        message += 'ログシートを新規作成しました。\n';
      } catch (e) {
        isSuccess = false;
        Logger.logError('ログシート作成エラー: ' + e.message);
        message += 'ログシート作成中にエラーが発生しました: ' + e.message + '\n';
      }
    }
    
    // 現在のシートを出品データシートに切り替え（存在する場合）
    const currentListingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    if (currentListingSheet) {
      ss.setActiveSheet(currentListingSheet);
    }
    
    // 完了ログを記録
    if (isSuccess) {
      Logger.endProcess('データ初期化が正常に完了しました');
      return {
        success: true,
        message: 'データ初期化が完了しました。\n' + message
      };
    } else {
      Logger.endProcess('データ初期化中に一部エラーが発生しました');
      return {
        success: false,
        message: '初期化中に一部エラーが発生しました。\n' + message
      };
    }
  } catch (error) {
    Logger.logError('データ初期化中に予期せぬエラーが発生: ' + error.message);
    return {
      success: false,
      message: '初期化中に予期せぬエラーが発生しました: ' + error.message
    };
  }
}

/**
 * サイドバーを再読み込みする
 * サイドバーからの呼び出し用
 */
function reloadSidebar() {
  try {
    UI.showSidebar();
    return { success: true };
  } catch (error) {
    Logger.logError('サイドバー再読み込みエラー: ' + error.message);
    return { 
      success: false,
      message: 'サイドバー再読み込みエラー: ' + error.message
    };
  }
} 