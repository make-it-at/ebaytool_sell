/**
 * eBay出品作業効率化ツール - フィルターモジュール
 * 
 * 各種フィルタリング機能を提供します。
 * 
 * バージョン: v1.2.0
 * 最終更新日: 2024-07-16
 */

// Filters名前空間
const Filters = {};

/**
 * エディタから直接実行するための所在地情報修正のグローバルエントリーポイント
 */
function runLocationFixFromEditor() {
  return Filters.runLocationFix();
}

/**
 * エディタから直接実行するためのNGワードフィルタリングのグローバルエントリーポイント
 */
function runNgWordFilterFromEditor() {
  return Filters.runNgWordFilter();
}

/**
 * エディタから直接実行するための重複チェックのグローバルエントリーポイント
 */
function runDuplicateCheckFromEditor() {
  return Filters.runDuplicateCheck();
}

/**
 * エディタから直接実行するための文字数制限フィルターのグローバルエントリーポイント
 */
function runLengthFilterFromEditor() {
  return Filters.runLengthFilter();
}

/**
 * エディタから直接実行するための価格フィルターのグローバルエントリーポイント
 */
function runPriceFilterFromEditor() {
  return Filters.runPriceFilter();
}

/**
 * NGワードフィルタリングを実行する
 */
Filters.runNgWordFilter = function() {
  Logger.startProcess('NGワードフィルタリング');
  UI.showProgressBar('NGワードフィルタリングを実行中...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // 出品データシートが存在するか確認
    if (!listingSheet) {
      throw new Error('出品データシートが見つかりません。初期設定を実行するか、データをインポートしてください。');
    }
    
    // データ範囲を取得
    const dataRange = listingSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行をスキップ
    const headerRow = values[0];
    const dataRows = values.slice(1);
    
    // Title列のインデックスを取得
    const titleColumnIndex = headerRow.indexOf('Title');
    if (titleColumnIndex === -1) {
      throw new Error('Title列が見つかりません。ヘッダー行に「Title」が含まれているか確認してください。');
    }
    
    // 処理結果列のインデックスを取得（なければ追加）
    const resultColumnIndex = headerRow.indexOf('処理結果');
    
    // 設定を取得
    const settings = Config.getSettings();
    const ngWords = settings.ngWords;
    const ngWordMode = settings.ngWordMode;
    
    // 設定内容をログに出力（デバッグ用）
    Logger.log(`NGワード設定: ${ngWords ? ngWords.length : 0}件のNGワード, モード: ${ngWordMode}`);
    if (ngWords && ngWords.length > 0) {
      ngWords.forEach(word => Logger.log(`NGワード: "${word}"`));
    }
    
    // 新しい結果データの準備
    let resultData = [];
    let rowsToDelete = [];
    
    // NGワードフィルタリング処理
    dataRows.forEach((row, index) => {
      // 処理の進捗状況を更新（10%単位）
      if (index % Math.floor(Math.max(dataRows.length, 10) / 10) === 0) {
        UI.updateProgressBar(Math.floor((index / Math.max(dataRows.length, 1)) * 100));
      }
      
      const title = row[titleColumnIndex]; // Title列の値
      
      // NGワードのチェック
      let containsNgWord = false;
      let processedTitle = title;
      let matchedNgWords = [];
      
      for (const ngWord of ngWords) {
        if (ngWord && title.toLowerCase().includes(ngWord.toLowerCase())) {
          containsNgWord = true;
          matchedNgWords.push(ngWord);
          
          // 部分削除モードの場合は、NGワードのみを削除
          if (ngWordMode === '部分削除モード') {
            processedTitle = processedTitle.replace(new RegExp(ngWord, 'gi'), '');
          } else {
            // リスト全削除モードはこの行を削除対象とするのでbreak
            break;
          }
        }
      }
      
      // リスト全削除モードでNGワードを含む場合は削除対象に追加
      if (containsNgWord && ngWordMode !== '部分削除モード') {
        rowsToDelete.push(index + 2); // +2 は1-indexedと、ヘッダー行をスキップするため
        Logger.log(`NGワード含有のためスキップ: "${title}", 一致NGワード: ${matchedNgWords.join(', ')}`);
      } else {
        // それ以外の場合は結果データに追加
        const newRow = [...row];
        
        // 部分削除モードの場合、タイトルを置き換え
        if (ngWordMode === '部分削除モード' && containsNgWord) {
          newRow[titleColumnIndex] = processedTitle;
          Logger.log(`NGワード部分削除: "${title}" → "${processedTitle}", 一致NGワード: ${matchedNgWords.join(', ')}`);
          
          // 処理結果列がある場合は更新
          if (resultColumnIndex >= 0) {
            newRow[resultColumnIndex] = 'NGワード部分削除';
          } else {
            newRow.push('NGワード部分削除');
          }
        } else {
          // 処理結果列がある場合は更新
          if (resultColumnIndex >= 0) {
            newRow[resultColumnIndex] = 'OK';
          } else {
            newRow.push('OK');
          }
        }
        
        resultData.push([index + 2, newRow]); // 行番号と新しい行データを保存
      }
    });
    
    // 処理結果を反映
    // まず、削除対象の行を削除（後ろから処理して行ずれを防止）
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a); // 降順にソート
      for (const rowIndex of rowsToDelete) {
        listingSheet.deleteRow(rowIndex);
      }
      
      // 削除後に結果データの行番号を再計算（削除した行より下の行は上にシフトする）
      resultData = resultData.map(([rowIndex, row]) => {
        let newRowIndex = rowIndex;
        for (const deletedRow of rowsToDelete) {
          if (deletedRow < rowIndex) {
            newRowIndex--;
          }
        }
        return [newRowIndex, row];
      });
    }
    
    // 更新データを反映（行ごとに更新）
    resultData.forEach(([rowIndex, row]) => {
      // 行が存在する場合のみ更新
      if (rowIndex >= 2 && rowIndex <= listingSheet.getLastRow()) {
        listingSheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      }
    });
    
    UI.showSuccessMessage(`NGワードフィルタリングが完了しました。処理結果: ${resultData.length}件 / ${dataRows.length}件 (${rowsToDelete.length}件削除)`);
    Logger.endProcess('NGワードフィルタリング完了');
    
    return true;
  } catch (error) {
    Logger.logError('NGワードフィルタリング中にエラー: ' + error.message);
    UI.showErrorMessage('NGワードフィルタリング中にエラーが発生しました: ' + error.message);
    return false;
  } finally {
    UI.hideProgressBar();
  }
};

/**
 * 重複チェックを実行する
 */
Filters.runDuplicateCheck = function() {
  Logger.startProcess('重複チェック');
  UI.showProgressBar('重複チェックを実行中...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // 出品データシートが存在するか確認
    if (!listingSheet) {
      throw new Error('出品データシートが見つかりません。初期設定を実行するか、データをインポートしてください。');
    }
    
    // データ範囲を取得
    const dataRange = listingSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行をスキップ
    const headerRow = values[0];
    const dataRows = values.slice(1);
    
    // 処理結果列のインデックスを取得（なければ追加）
    const resultColumnIndex = headerRow.indexOf('処理結果');
    
    // 設定を取得
    const settings = Config.getSettings();
    const duplicateThreshold = settings.duplicateThreshold;
    
    // 重複チェック結果の準備
    let rowsToDelete = [];
    
    // タイトルごとの類似度計算
    for (let i = 0; i < dataRows.length; i++) {
      // 処理の進捗状況を更新（10%単位）
      if (i % Math.floor(Math.max(dataRows.length, 10) / 10) === 0) {
        UI.updateProgressBar(Math.floor((i / Math.max(dataRows.length, 1)) * 100));
      }
      
      // すでに削除対象として記録されている行はスキップ
      if (rowsToDelete.includes(i + 2)) continue;
      
      const title1 = dataRows[i][0]; // 商品名
      
      for (let j = i + 1; j < dataRows.length; j++) {
        if (rowsToDelete.includes(j + 2)) continue;
        
        const title2 = dataRows[j][0]; // 比較対象の商品名
        
        // 類似度を計算（レーベンシュタイン距離を使用）
        const similarity = this.calculateSimilarity(title1, title2);
        
        // 閾値以上の類似度がある場合は重複と判定
        if (similarity >= duplicateThreshold) {
          rowsToDelete.push(j + 2); // +2 は1-indexedと、ヘッダー行をスキップするため
          Logger.log(`重複検出: "${title1}" と "${title2}" （類似度: ${similarity}%）`);
        }
      }
    }
    
    // 処理結果を反映
    // 削除対象の行を削除（後ろから処理して行ずれを防止）
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a); // 降順にソート
      for (const rowIndex of rowsToDelete) {
        listingSheet.deleteRow(rowIndex);
      }
    }
    
    // 残った行に処理結果を更新
    if (resultColumnIndex >= 0) {
      const lastRow = listingSheet.getLastRow();
      if (lastRow > 1) { // ヘッダー行より下に行が存在する場合
        const resultRange = listingSheet.getRange(2, resultColumnIndex + 1, lastRow - 1, 1);
        const currentValues = resultRange.getValues();
        
        // 各行の処理結果を「OK」または既存値+「重複チェック完了」に更新
        const newValues = currentValues.map(([value]) => {
          if (!value) return ['OK'];
          return [value + ', 重複チェック完了'];
        });
        
        resultRange.setValues(newValues);
      }
    }
    
    UI.showSuccessMessage(`重複チェックが完了しました。${rowsToDelete.length}件の重複を除外しました。`);
    Logger.endProcess('重複チェック完了');
    
    return true;
  } catch (error) {
    Logger.logError('重複チェック中にエラー: ' + error.message);
    UI.showErrorMessage('重複チェック中にエラーが発生しました: ' + error.message);
    return false;
  } finally {
    UI.hideProgressBar();
  }
};

/**
 * 文字列の類似度を計算する（0-100%）
 */
Filters.calculateSimilarity = function(str1, str2) {
  // レーベンシュタイン距離を計算
  const levDistance = this.levenshteinDistance(str1.toLowerCase(), str2.toLowerCase());
  
  // 最大長に対する距離の比率から類似度を計算
  const maxLength = Math.max(str1.length, str2.length);
  const similarity = ((maxLength - levDistance) / maxLength) * 100;
  
  return Math.round(similarity);
};

/**
 * レーベンシュタイン距離を計算する
 */
Filters.levenshteinDistance = function(a, b) {
  const matrix = [];
  
  // 行列の初期化
  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }
  
  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }
  
  // 距離を計算
  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // 置換
          matrix[i][j - 1] + 1,     // 挿入
          matrix[i - 1][j] + 1      // 削除
        );
      }
    }
  }
  
  return matrix[b.length][a.length];
};

/**
 * 文字数制限フィルターを実行する
 */
Filters.runLengthFilter = function() {
  Logger.startProcess('文字数制限フィルタリング');
  UI.showProgressBar('文字数制限フィルタリングを実行中...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // 出品データシートが存在するか確認
    if (!listingSheet) {
      throw new Error('出品データシートが見つかりません。初期設定を実行するか、データをインポートしてください。');
    }
    
    // データ範囲を取得
    const dataRange = listingSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行をスキップ
    const headerRow = values[0];
    const dataRows = values.slice(1);
    
    // 処理結果列のインデックスを取得（なければ追加）
    const resultColumnIndex = headerRow.indexOf('処理結果');
    
    // 設定を取得
    const settings = Config.getSettings();
    const characterLimit = settings.characterLimit;
    
    // 削除対象の行を特定
    let rowsToDelete = [];
    
    // フィルタリング処理
    dataRows.forEach((row, index) => {
      // 処理の進捗状況を更新（10%単位）
      if (index % Math.floor(Math.max(dataRows.length, 10) / 10) === 0) {
        UI.updateProgressBar(Math.floor((index / Math.max(dataRows.length, 1)) * 100));
      }
      
      const title = row[0]; // 商品名
      
      // 文字数チェック
      if (title.length < characterLimit) {
        rowsToDelete.push(index + 2); // +2 は1-indexedと、ヘッダー行をスキップするため
        Logger.log(`文字数不足のためスキップ: "${title}" (${title.length}文字)`);
      }
    });
    
    // 処理結果を反映
    // 削除対象の行を削除（後ろから処理して行ずれを防止）
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a); // 降順にソート
      for (const rowIndex of rowsToDelete) {
        listingSheet.deleteRow(rowIndex);
      }
    }
    
    // 残った行に処理結果を更新
    if (resultColumnIndex >= 0) {
      const lastRow = listingSheet.getLastRow();
      if (lastRow > 1) { // ヘッダー行より下に行が存在する場合
        const resultRange = listingSheet.getRange(2, resultColumnIndex + 1, lastRow - 1, 1);
        const currentValues = resultRange.getValues();
        
        // 各行の処理結果を「OK」または既存値+「文字数OK」に更新
        const newValues = currentValues.map(([value]) => {
          if (!value) return ['OK'];
          return [value + ', 文字数OK'];
        });
        
        resultRange.setValues(newValues);
      }
    }
    
    UI.showSuccessMessage(`文字数制限フィルタリングが完了しました。${rowsToDelete.length}件が文字数不足でスキップされました。`);
    Logger.endProcess('文字数制限フィルタリング完了');
    
    return true;
  } catch (error) {
    Logger.logError('文字数制限フィルタリング中にエラー: ' + error.message);
    UI.showErrorMessage('文字数制限フィルタリング中にエラーが発生しました: ' + error.message);
    return false;
  } finally {
    UI.hideProgressBar();
  }
};

/**
 * 所在地情報修正を実行する
 */
Filters.runLocationFix = function() {
  Logger.startProcess('所在地情報修正');
  UI.showProgressBar('所在地情報を修正中...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // 出品データシートが存在するか確認
    if (!listingSheet) {
      throw new Error('出品データシートが見つかりません。初期設定を実行するか、データをインポートしてください。');
    }
    
    // データ範囲を取得
    const dataRange = listingSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行をスキップ
    const headerRow = values[0];
    const dataRows = values.slice(1);
    
    // 処理結果列のインデックスを取得（なければ追加）
    const resultColumnIndex = headerRow.indexOf('処理結果');
    
    // Locationカラムのインデックスを取得（ヘッダーから位置を特定）
    let locationColumnIndex = headerRow.indexOf('所在地');
    
    // Locationカラムが見つからない場合はLocationも試す
    if (locationColumnIndex === -1) {
      const altLocationColumnIndex = headerRow.indexOf('Location');
      if (altLocationColumnIndex === -1) {
        throw new Error('所在地列が見つかりません。ヘッダー行に「所在地」または「Location」が含まれているか確認してください。');
      } else {
        // Locationカラムが見つかった場合
        Logger.log('「Location」列を所在地情報として使用します');
        locationColumnIndex = altLocationColumnIndex;
      }
    }
    
    // 進捗表示のために更新
    UI.updateProgressBar(10);
    
    // 結果データの準備
    let updatedLocations = [];
    
    // 所在地修正処理
    dataRows.forEach((row, index) => {
      // 処理の進捗状況を更新（10%単位）
      if (index % Math.floor(Math.max(dataRows.length, 10) / 10) === 0) {
        UI.updateProgressBar(10 + Math.floor((index / Math.max(dataRows.length, 1)) * 80));
      }
      
      let location = row[locationColumnIndex]; // 所在地カラムの値
      let originalLocation = location;
      
      // 数字を削除するシンプルな処理
      try {
        if (location && typeof location === 'string') {
          location = location.replace(/[0-9]+/g, '');
        }
      } catch (e) {
        // エラーが発生しても処理を続行
        Logger.logError('所在地情報の数字削除中にエラー（スキップして続行）: ' + e.message);
      }
      
      // 所在地が変更された場合にのみ更新リストに追加
      if (location !== originalLocation) {
        updatedLocations.push({
          row: index + 2, // +2 は1-indexedと、ヘッダー行をスキップするため
          column: locationColumnIndex + 1, // +1 は0-indexedから1-indexedに変換
          value: location
        });
      }
    });
    
    // 進捗表示を更新
    UI.updateProgressBar(90);
    
    // 処理結果を反映
    // 所在地情報を更新
    updatedLocations.forEach((update, index) => {
      listingSheet.getRange(update.row, update.column).setValue(update.value);
      
      // 進捗表示の細かい更新
      if (index % Math.floor(Math.max(updatedLocations.length, 10) / 10) === 0) {
        UI.updateProgressBar(90 + Math.floor((index / Math.max(updatedLocations.length, 1)) * 10));
      }
    });
    
    // 処理結果列を更新
    if (resultColumnIndex >= 0) {
      const lastRow = listingSheet.getLastRow();
      if (lastRow > 1) { // ヘッダー行より下に行が存在する場合
        const resultRange = listingSheet.getRange(2, resultColumnIndex + 1, lastRow - 1, 1);
        const currentValues = resultRange.getValues();
        
        // 各行の処理結果を「OK」または既存値+「所在地修正完了」に更新
        const newValues = currentValues.map(([value]) => {
          if (!value) return ['OK'];
          return [value + ', 所在地修正完了'];
        });
        
        resultRange.setValues(newValues);
      }
    }
    
    UI.updateProgressBar(100);
    
    UI.showSuccessMessage(`所在地情報の修正が完了しました。${updatedLocations.length}件の所在地情報を修正しました。`);
    Logger.endProcess('所在地情報修正完了');
    
    return true;
  } catch (error) {
    Logger.logError('所在地情報修正中にエラー: ' + error.message);
    UI.showErrorMessage('所在地情報修正中にエラーが発生しました: ' + error.message);
    return false;
  } finally {
    UI.hideProgressBar();
  }
};

/**
 * 価格フィルタリングを実行する
 */
Filters.runPriceFilter = function() {
  Logger.startProcess('価格フィルタリング');
  UI.showProgressBar('価格フィルタリングを実行中...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listingSheet = ss.getSheetByName(Config.SHEET_NAMES.LISTING);
    
    // 出品データシートが存在するか確認
    if (!listingSheet) {
      throw new Error('出品データシートが見つかりません。初期設定を実行するか、データをインポートしてください。');
    }
    
    // データ範囲を取得
    const dataRange = listingSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行をスキップ
    const headerRow = values[0];
    const dataRows = values.slice(1);
    
    // 処理結果列のインデックスを取得（なければ追加）
    const resultColumnIndex = headerRow.indexOf('処理結果');
    
    // 設定を取得
    const settings = Config.getSettings();
    const priceThreshold = settings.priceThreshold;
    
    // 削除対象の行を特定
    let rowsToDelete = [];
    
    // 価格フィルタリング処理
    dataRows.forEach((row, index) => {
      // 処理の進捗状況を更新（10%単位）
      if (index % Math.floor(Math.max(dataRows.length, 10) / 10) === 0) {
        UI.updateProgressBar(Math.floor((index / Math.max(dataRows.length, 1)) * 100));
      }
      
      const price = parseFloat(row[1]); // 価格($)
      
      // 価格チェック
      if (isNaN(price) || price <= priceThreshold) {
        rowsToDelete.push(index + 2); // +2 は1-indexedと、ヘッダー行をスキップするため
        Logger.log(`価格条件不一致のためスキップ: $${price}`);
      }
    });
    
    // 処理結果を反映
    // 削除対象の行を削除（後ろから処理して行ずれを防止）
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a); // 降順にソート
      for (const rowIndex of rowsToDelete) {
        listingSheet.deleteRow(rowIndex);
      }
    }
    
    // 残った行に処理結果を更新
    if (resultColumnIndex >= 0) {
      const lastRow = listingSheet.getLastRow();
      if (lastRow > 1) { // ヘッダー行より下に行が存在する場合
        const resultRange = listingSheet.getRange(2, resultColumnIndex + 1, lastRow - 1, 1);
        const currentValues = resultRange.getValues();
        
        // 各行の処理結果を「OK」または既存値+「価格OK」に更新
        const newValues = currentValues.map(([value]) => {
          if (!value) return ['OK'];
          return [value + ', 価格OK'];
        });
        
        resultRange.setValues(newValues);
      }
    }
    
    UI.showSuccessMessage(`価格フィルタリングが完了しました。${rowsToDelete.length}件が価格条件で除外されました。`);
    Logger.endProcess('価格フィルタリング完了');
    
    return true;
  } catch (error) {
    Logger.logError('価格フィルタリング中にエラー: ' + error.message);
    UI.showErrorMessage('価格フィルタリング中にエラーが発生しました: ' + error.message);
    return false;
  } finally {
    UI.hideProgressBar();
  }
}; 