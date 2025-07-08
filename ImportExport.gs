/**
 * eBay出品作業効率化ツール - インポート/エクスポートモジュール
 * 
 * CSVのインポートとエクスポート機能を提供します。
 * 
 * バージョン: v1.5.3
 * 最終更新日: 2025-06-15
 * 更新内容: データ処理の高速化（バッチ処理の実装とメモリ使用量の最適化）
 */

// ImportExport名前空間
const ImportExport = {};

/**
 * CSVファイルをインポートする
 * @param {Blob} csvFile インポートするCSVファイル
 * @return {boolean} 成功したかどうか
 */
ImportExport.importCsv = function(csvFile) {
  Logger.startProcess('CSVインポート');
  UI.showProgressBar('CSVファイルをインポートしています...');
  
  try {
    // CSVファイルを読み込み
    const csvString = csvFile.getDataAsString();
    const csvData = Utilities.parseCsv(csvString);
    
    // 行数チェック
    if (csvData.length > Config.MAX_ROWS) {
      throw new Error(`インポートするデータが多すぎます。最大${Config.MAX_ROWS}行までです。`);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // シートが存在しない場合は作成
    if (!listingSheet) {
      listingSheet = ss.insertSheet(Config.SHEET_NAMES.LISTING);
      
      // ヘッダー行を設定（必要に応じて）
      if (Config.SHEET_HEADERS && Config.SHEET_HEADERS.LISTING) {
        listingSheet.getRange(1, 1, 1, Config.SHEET_HEADERS.LISTING.length)
          .setValues([Config.SHEET_HEADERS.LISTING]);
        listingSheet.setFrozenRows(1);
      }
      
      Logger.logInfo('出品データシートが見つからなかったため、新規作成しました。');
    }
    
    // シートをクリア
    if (listingSheet.getLastRow() > 0) {
      listingSheet.clearContents();
    }
    
    // CSVデータをバッチで処理して書き込み
    if (csvData.length > 0) {
      const totalRows = csvData.length;
      const batchSize = 1000; // バッチサイズを1000行に設定
      const columnCount = csvData[0].length;
      
      // 進捗更新のための変数
      let processedRows = 0;
      
      // バッチ処理
      for (let i = 0; i < totalRows; i += batchSize) {
        // 現在のバッチサイズを計算（最後のバッチは小さくなる可能性あり）
        const currentBatchSize = Math.min(batchSize, totalRows - i);
        
        // 現在のバッチのデータを取得
        const batchData = csvData.slice(i, i + currentBatchSize);
        
        // バッチデータをシートに書き込み
        listingSheet.getRange(i + 1, 1, currentBatchSize, columnCount).setValues(batchData);
      
        // 進捗を更新
        processedRows += currentBatchSize;
        UI.updateProgressBar(Math.floor((processedRows / totalRows) * 100));
        
        // メモリを解放する時間を確保
        if (i + batchSize < totalRows) {
          Utilities.sleep(50); // 次のバッチの前に短い遅延
        }
      }
      
      // 不要になった大きな変数を明示的に解放
      csvData.length = 0;
    }
    
    // 処理終了メッセージ
    UI.showSuccessMessage(`CSVインポートが完了しました。${csvData.length}件のデータをインポートしました。`);
    Logger.endProcess('CSVインポート完了');
    
    return true;
  } catch (error) {
    Logger.logError('CSVインポート中にエラーが発生: ' + error.message);
    UI.showErrorMessage('CSVインポート中にエラーが発生しました: ' + error.message);
    return false;
  } finally {
    UI.hideProgressBar();
  }
};

/**
 * ヘッダーのマッピングを作成する
 * @param {Array} sourceHeaders ソースのヘッダー配列
 * @param {Array} targetHeaders ターゲットのヘッダー配列
 * @return {Object} インデックスのマッピング
 */
ImportExport.createHeaderMapping = function(sourceHeaders, targetHeaders) {
  const mapping = [];
  
  // ヘッダーインデックスのキャッシュを作成（検索を高速化）
  const headerIndexMap = {};
  sourceHeaders.forEach((header, index) => {
    headerIndexMap[header.toLowerCase()] = index;
  });
  
  // 各ターゲットヘッダーに対応するソースヘッダーのインデックスを検索
  for (let i = 0; i < targetHeaders.length; i++) {
    const targetHeader = targetHeaders[i].toLowerCase();
    let sourceIndex = -1;
    
    // 完全一致をキャッシュから探す（O(1)の操作）
    if (headerIndexMap[targetHeader] !== undefined) {
      sourceIndex = headerIndexMap[targetHeader];
    } else {
      // 部分一致の場合はループで探す（最適化の余地が少ない）
      for (let j = 0; j < sourceHeaders.length; j++) {
        const sourceHeader = sourceHeaders[j].toLowerCase();
        if (sourceHeader.includes(targetHeader) || targetHeader.includes(sourceHeader)) {
          sourceIndex = j;
          break;
        }
      }
    }
    
    mapping.push(sourceIndex);
  }
  
  return mapping;
};

/**
 * 出品データシートのデータをeBay形式にフォーマットしてCSVエクスポート
 * @return {Array} eBay形式のデータ（ヘッダー行を含む）
 */
ImportExport.formatForEbay = function() {
  Logger.startProcess('eBay形式フォーマット');
  UI.showProgressBar('eBayフォーマットに変換しています...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // データ範囲を取得（一度に全てのデータを取得する代わりにチャンク処理）
    const totalRows = listingSheet.getLastRow();
    const totalColumns = listingSheet.getLastColumn();
    
    if (totalRows <= 1) {
      throw new Error('出品データが見つかりません。先にデータをインポートしてください。');
    }
    
    // ヘッダー行を取得
    const headerRow = listingSheet.getRange(1, 1, 1, totalColumns).getValues()[0];
    
    // eBay形式のデータを準備
    const ebayData = [];
    
    // ヘッダー行を追加
    ebayData.push(Config.SHEET_HEADERS.EBAY_FORMAT);
    
    // バッチサイズを設定
    const batchSize = 500;
    const dataRows = totalRows - 1; // ヘッダーを除く
    
    // コンディションマップをキャッシュ（毎回の関数呼び出しを減らす）
    const conditionMapCache = this.createConditionMapCache();
    
    // バッチ処理でデータを変換
    for (let i = 0; i < dataRows; i += batchSize) {
      // 現在のバッチサイズを計算
      const currentBatchSize = Math.min(batchSize, dataRows - i);
      
      // データバッチを取得
      const rowsData = listingSheet.getRange(i + 2, 1, currentBatchSize, totalColumns).getValues();
      
      // 各行を処理
      rowsData.forEach((row, index) => {
        // eBay形式の行を作成（キャッシュしたコンディションマップを使用）
        const ebayRow = this.createEbayRowOptimized(row, conditionMapCache);
      ebayData.push(ebayRow);
        
        // 進捗を更新（10%ごと）
        if (index % Math.floor(Math.max(currentBatchSize, 10) / 10) === 0) {
          const overallProgress = ((i + index) / dataRows) * 100;
          UI.updateProgressBar(Math.floor(overallProgress));
        }
      });
      
      // メモリを解放する時間を確保
      if (i + batchSize < dataRows) {
        Utilities.sleep(50);
      }
    }
    
    UI.showSuccessMessage(`eBayフォーマットへの変換が完了しました。${dataRows}件のデータを変換しました。`);
    Logger.endProcess('eBayフォーマット変換完了');
    
    return ebayData;
  } catch (error) {
    Logger.logError('eBayフォーマット変換中にエラー: ' + error.message);
    UI.showErrorMessage('eBayフォーマット変換中にエラーが発生しました: ' + error.message);
    return null;
  } finally {
    UI.hideProgressBar();
  }
};

/**
 * コンディションマップのキャッシュを作成
 * @return {Object} コンディションマップのキャッシュ
 */
ImportExport.createConditionMapCache = function() {
  return {
    'new': '1000', // New
    'used': '3000', // Used
    'like new': '1500', // Like New
    'very good': '2000', // Very Good
    'good': '2500', // Good
    'acceptable': '3500', // Acceptable
    'for parts or not working': '7000' // For parts or not working
  };
};

/**
 * 最適化されたeBay形式の行データを作成する
 * @param {Array} row 元の行データ
 * @param {Object} conditionMapCache コンディションマップのキャッシュ
 * @return {Array} eBay形式の行データ
 */
ImportExport.createEbayRowOptimized = function(row, conditionMapCache) {
  // 元データの項目
  const title = row[0]; // 商品名
  const price = row[1]; // 価格($)
  const location = row[2]; // 所在地
  const condition = row[3]; // コンディション
  
  // eBay形式の行データを作成
  const ebayRow = new Array(Config.SHEET_HEADERS.EBAY_FORMAT.length).fill('');
  
  // 必須項目を設定
  ebayRow[0] = 'Add'; // Action
  ebayRow[2] = title; // Title
  
  // コンディションのマッピング（キャッシュを使用）
  const conditionLower = condition ? condition.toLowerCase() : '';
  ebayRow[8] = conditionMapCache[conditionLower] || '3000'; // デフォルトは'Used'
  
  ebayRow[11] = 'FixedPrice'; // Format
  ebayRow[12] = price; // Start price
  ebayRow[14] = '1'; // Quantity
  ebayRow[15] = '1'; // PayPal accepted
  ebayRow[17] = '1'; // Immediate payment required
  ebayRow[18] = location; // Location
  ebayRow[19] = 'USPSFirstClass'; // Shipping service 1
  ebayRow[20] = '0'; // Shipping service cost 1
  ebayRow[25] = '3'; // Max dispatch time
  ebayRow[26] = '1'; // Returns accepted
  ebayRow[27] = 'MoneyBack'; // Refund
  ebayRow[28] = 'Seller'; // Return shipping cost paid by
  
  return ebayRow;
};

/**
 * コンディション文字列をeBayのコンディションコードにマッピングする
 * @param {string} conditionText コンディションの文字列
 * @return {string} eBayのコンディションコード
 */
ImportExport.mapCondition = function(conditionText) {
  const conditionMap = {
    'new': '1000', // New
    'used': '3000', // Used
    'like new': '1500', // New other
    'good': '4000', // Very Good
    'acceptable': '6000', // Acceptable
    'for parts': '7000' // For parts or not working
  };
  
  // nullやundefinedまたは文字列でない場合は処理せずデフォルト値を返す
  if (conditionText === null || conditionText === undefined || typeof conditionText !== 'string') {
    return '3000'; // デフォルトはUsed
  }
  
  // 小文字に変換して検索
  const lowerCondition = conditionText.toLowerCase();
  
  // マッピングを検索
  for (const key in conditionMap) {
    if (lowerCondition.includes(key)) {
      return conditionMap[key];
    }
  }
  
  // デフォルトはUsed
  return '3000';
};

/**
 * 出品データシートの内容をそのままCSVでエクスポートする
 * eBay形式に変換せず、シートの内容をそのままエクスポートする
 * @param {string} sheetName エクスポート対象のシート名（デフォルト: 出品データ）
 * @return {Object} エクスポート結果（csvData: CSV内容, fileName: ファイル名）
 */
ImportExport.exportToCsv = function(sheetName) {
  Logger.startProcess('CSVエクスポート');
  UI.showProgressBar('CSVエクスポートを実行中...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // シート名が指定されていない場合は出品データシートをデフォルトとする
    if (!sheetName) {
      sheetName = Config.SHEET_NAMES.LISTING;
    }
    
    // 指定されたシートを取得
    const sheet = ss.getSheetByName(sheetName);
    
    // シートが存在しない場合はエラー
    if (!sheet) {
      throw new Error(`指定されたシート "${sheetName}" が見つかりません。`);
    }
    
    // データ範囲を取得
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    UI.updateProgressBar(30);
    
    // データがない場合はメッセージを表示して終了
    if (!values || values.length <= 1) { // ヘッダーのみ以下の場合
      UI.updateProgressBar(100);
      UI.showErrorMessage('データがありません。まずデータをインポートしてください。');
      Logger.endProcess('CSVエクスポート中止 - データなし');
      return null;
    }
    
    UI.updateProgressBar(50);
    
    // CSVデータの作成
    let csvData = '';
    values.forEach((row, index) => {
      // 進捗状況を更新（50%〜90%）
      if (index % Math.floor(Math.max(values.length, 10) / 10) === 0) {
        const progress = 50 + Math.floor((index / Math.max(values.length, 1)) * 40);
        UI.updateProgressBar(progress);
      }
      
      // 各セルをCSV形式に変換
      const csvRow = row.map(cell => {
        // nullやundefinedを空文字に変換
        if (cell === null || cell === undefined) {
          return '';
        }
        
        // 文字列に変換
        let cellStr = cell.toString();
        
        // カンマやダブルクォートを含む場合はダブルクォートで囲み、内部のダブルクォートはエスケープ
        if (cellStr.includes(',') || cellStr.includes('"')) {
          return '"' + cellStr.replace(/"/g, '""') + '"';
        }
        return cellStr;
      }).join(',');
      
      csvData += csvRow + '\r\n';
    });
    
    UI.updateProgressBar(90);
    
    // シート名をファイル名に含める
    const now = new Date();
    const timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    const fileName = `ebay_${sheetName}_${timestamp}.csv`;
    
    // データ行数を計算（ヘッダー行を除く）
    const rowCount = Math.max(0, values.length - 1);
    const message = `CSVエクスポートが完了しました。シート "${sheetName}" から${rowCount}件のデータをエクスポートしました。`;
    
    UI.updateProgressBar(100);
    UI.showSuccessMessage(message);
    Logger.endProcess('CSVエクスポート完了');
    
    // CSVデータとファイル名を返す
    return {
      csvData: csvData,
      fileName: fileName
    };
  } catch (error) {
    Logger.logError('CSVエクスポート中にエラー: ' + error.message);
    UI.showErrorMessage('CSVエクスポート中にエラーが発生しました: ' + error.message);
    return null;
  } finally {
    UI.hideProgressBar();
  }
}; 