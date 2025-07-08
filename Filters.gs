/**
 * eBay出品作業効率化ツール - フィルターモジュール
 * 
 * 各種フィルタリング機能を提供します。
 * 
 * バージョン: v1.5.3
 * 最終更新日: 2025-06-15
 * 更新内容: NGワードフィルター処理の高速化（バッチ処理の実装）
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
 * @return {Object} 処理結果
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
    
    // データの基本情報を取得
    const lastRow = listingSheet.getLastRow();
    const lastColumn = listingSheet.getLastColumn();
    
    if (lastRow <= 1) {
      throw new Error('出品データが見つかりません。データをインポートしてください。');
    }
    
    // ヘッダー行のみ取得
    const headerRow = listingSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    
    // Title列のインデックスを取得
    const titleColumnIndex = headerRow.indexOf('Title');
    if (titleColumnIndex === -1) {
      throw new Error('Title列が見つかりません。ヘッダー行に「Title」が含まれているか確認してください。');
    }
    
    // 設定を取得
    const settings = Config.getSettings();
    const deleteListNgWords = settings.deleteListNgWords || [];
    const deletePartNgWords = settings.deletePartNgWords || [];
    
    // 検索用に正規化されたNGワードリストを作成
    // 大文字・小文字の区別をなくし、連続する空白を単一の空白に置換
    const normalizedListNgWords = deleteListNgWords.map(word => 
      this.normalizeSearchTerm(word)
    );
    
    const normalizedPartNgWords = deletePartNgWords.map(word => 
      this.normalizeSearchTerm(word)
    );
    
    // 設定内容をログに出力（デバッグ用）
    Logger.log(`NGワード設定: リスト削除=${deleteListNgWords.length}件, 部分削除=${deletePartNgWords.length}件`);
    
    // 処理前のデータ行数（ヘッダーを除く）
    const beforeDataCount = lastRow - 1;
    
    // バッチ処理のためのパラメータ
    const batchSize = 500; // 一度に処理する行数
    let rowsToDelete = []; // 削除対象の行
    let rowsToUpdate = []; // 更新対象の行と内容
    
    // データをバッチで処理
    const dataRows = lastRow - 1; // ヘッダーを除いた行数
    
    // NGワード削除用の正規表現パターンをキャッシュ
    const partNgWordRegexCache = deletePartNgWords.map(ngWord => {
      const escapeRegExp = (string) => string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      return {
        pattern: new RegExp(escapeRegExp(ngWord).replace(/\s+/g, '\\s+'), 'gi'),
        original: ngWord
      };
    });
    
    for (let batchStart = 0; batchStart < dataRows; batchStart += batchSize) {
      // 現在のバッチサイズを計算
      const currentBatchSize = Math.min(batchSize, dataRows - batchStart);
      
      // バッチデータを取得（2行目からのデータを取得するため、行番号は+2）
      const batchData = listingSheet.getRange(batchStart + 2, 1, currentBatchSize, lastColumn).getValues();
    
      // バッチデータを処理
      for (let i = 0; i < batchData.length; i++) {
        const row = batchData[i];
        const rowIndex = batchStart + i + 2; // 実際のシートの行番号
        
        // 進捗状況を更新
        if (i % Math.floor(Math.max(currentBatchSize, 10) / 10) === 0) {
          const overallProgress = ((batchStart + i) / dataRows) * 100;
          UI.updateProgressBar(Math.floor(overallProgress));
      }
      
      const title = row[titleColumnIndex]; // Title列の値
      
      // タイトルを検索用に正規化
      const normalizedTitle = this.normalizeSearchTerm(title);
      
      // リスト削除NGワードのチェック
      let containsListNgWord = false;
      let matchedListNgWords = [];
      
        for (let j = 0; j < normalizedListNgWords.length; j++) {
          const ngWord = normalizedListNgWords[j];
        if (ngWord && normalizedTitle.includes(ngWord)) {
          containsListNgWord = true;
            matchedListNgWords.push(deleteListNgWords[j]); // 元の形式で記録
          break; // 1つでも見つかれば削除対象
        }
      }
      
      // リスト削除NGワードを含む場合は削除対象に追加
      if (containsListNgWord) {
          rowsToDelete.push(rowIndex); // 実際のシート行番号
        Logger.log(`リスト削除NGワード含有のためスキップ: "${title}", 一致NGワード: ${matchedListNgWords.join(', ')}`);
      } else {
        // それ以外の場合は部分削除NGワードのチェック
        let processedTitle = title;
        let matchedPartNgWords = [];
          let titleModified = false;
          
          // 高速化のためにキャッシュした正規表現を使用
          for (let j = 0; j < partNgWordRegexCache.length; j++) {
            const { pattern, original } = partNgWordRegexCache[j];
          
            if (this.normalizeSearchTerm(processedTitle).includes(normalizedPartNgWords[j])) {
              matchedPartNgWords.push(original);
              const oldTitle = processedTitle;
              processedTitle = processedTitle.replace(pattern, '');
              
              if (oldTitle !== processedTitle) {
                titleModified = true;
              }
            }
          }
        
        // 部分削除NGワードを処理した場合、タイトルを置き換え
          if (titleModified) {
            const newRow = [...row];
          newRow[titleColumnIndex] = processedTitle;
            rowsToUpdate.push({ row: rowIndex, data: newRow });
          Logger.log(`部分削除NGワード処理: "${title}" → "${processedTitle}", 一致NGワード: ${matchedPartNgWords.join(', ')}`);
          }
        }
        }
        
      // バッチ処理の間にメモリを解放
      if (batchStart + batchSize < dataRows) {
        Utilities.sleep(50);
      }
    }
    
    // 処理結果を反映
    // 削除対象の行を削除（後ろから処理して行ずれを防止）
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a); // 降順にソート
      
      // バッチでの削除処理（パフォーマンス向上のため）
      const deleteBatchSize = 50; // 削除のバッチサイズ
      for (let i = 0; i < rowsToDelete.length; i += deleteBatchSize) {
        const batch = rowsToDelete.slice(i, i + deleteBatchSize);
        for (const rowIndex of batch) {
        listingSheet.deleteRow(rowIndex);
      }
      
        // 削除処理の間にわずかな遅延を挿入
        if (i + deleteBatchSize < rowsToDelete.length) {
          Utilities.sleep(50);
        }
        
        // 削除進捗の更新
        UI.updateProgressBar(Math.floor(80 + (i / rowsToDelete.length) * 10));
      }
      
      // 削除した行に基づいて更新対象の行番号を調整
      if (rowsToUpdate.length > 0) {
        rowsToUpdate = rowsToUpdate.map(item => {
          let newRowIndex = item.row;
        for (const deletedRow of rowsToDelete) {
            if (deletedRow < newRowIndex) {
            newRowIndex--;
          }
        }
          return { row: newRowIndex, data: item.data };
      });
      }
    }
    
    // 更新データを反映（バッチで更新）
    if (rowsToUpdate.length > 0) {
      const updateBatchSize = 50; // 更新のバッチサイズ
      
      for (let i = 0; i < rowsToUpdate.length; i += updateBatchSize) {
        const batch = rowsToUpdate.slice(i, i + updateBatchSize);
        
        for (const item of batch) {
      // 行が存在する場合のみ更新
          if (item.row >= 2 && item.row <= listingSheet.getLastRow()) {
            listingSheet.getRange(item.row, 1, 1, item.data.length).setValues([item.data]);
      }
        }
        
        // 更新処理の間にわずかな遅延を挿入
        if (i + updateBatchSize < rowsToUpdate.length) {
          Utilities.sleep(50);
        }
        
        // 更新進捗の更新
        UI.updateProgressBar(Math.floor(90 + (i / rowsToUpdate.length) * 10));
      }
    }
    
    // 処理後のデータ行数を取得
    const afterDataCount = listingSheet.getLastRow() - 1;
    
    // 処理結果メッセージを作成
    const resultMessage = `NGワードフィルタリングが完了しました。データ数: ${beforeDataCount}件 → ${afterDataCount}件（削除: ${rowsToDelete.length}件, 修正: ${rowsToUpdate.length}件）`;
    UI.showSuccessMessage(resultMessage);
    
    // 統計情報を返す
    const stats = {
      beforeCount: beforeDataCount,
      afterCount: afterDataCount,
        removedCount: rowsToDelete.length,
      modifiedCount: rowsToUpdate.length
    };
    
    Logger.endProcess('NGワードフィルタリング完了');
    return { success: true, message: resultMessage, stats: stats };
  } catch (error) {
    Logger.logError('NGワードフィルタリング中にエラー: ' + error.message);
    UI.showErrorMessage('NGワードフィルタリング中にエラーが発生しました: ' + error.message);
    return { success: false, message: error.message };
  } finally {
    UI.hideProgressBar();
  }
};

/**
 * 検索用に文字列を正規化する
 * - 大文字・小文字を区別しないようにする（すべて小文字に変換）
 * - 連続する空白文字を1つの空白に置換
 * - 先頭・末尾の空白を削除
 * 
 * @param {string} text 正規化する文字列
 * @return {string} 正規化された文字列
 */
Filters.normalizeSearchTerm = function(text) {
  if (!text) return '';
  
  return text
    .toString()
    .toLowerCase()          // 小文字に変換
    .replace(/\s+/g, ' ')   // 連続する空白を1つの空白に置換
    .trim();                // 先頭・末尾の空白を削除
};

/**
 * 重複チェックを実行する
 * @return {Object} 処理結果
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
    
    // Title列のインデックスを取得
    const titleColumnIndex = headerRow.indexOf('Title');
    if (titleColumnIndex === -1) {
      throw new Error('Title列が見つかりません。ヘッダー行に「Title」が含まれているか確認してください。');
    }
    
    // 重複チェック結果の準備
    let rowsToDelete = [];
    let uniqueTitles = new Set();
    
    // 完全一致の重複チェック
    dataRows.forEach((row, index) => {
      // 処理の進捗状況を更新（10%単位）
      if (index % Math.floor(Math.max(dataRows.length, 10) / 10) === 0) {
        UI.updateProgressBar(Math.floor((index / Math.max(dataRows.length, 1)) * 100));
      }
      
      const title = row[titleColumnIndex]; // Title列の値
      
      // 完全一致の重複チェック
      if (uniqueTitles.has(title)) {
        rowsToDelete.push(index + 2); // +2 は1-indexedと、ヘッダー行をスキップするため
        Logger.log(`完全一致の重複を検出: "${title}"`);
      } else {
        uniqueTitles.add(title);
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
    
    // 結果メッセージ
    const message = `重複チェックが完了しました。${rowsToDelete.length}件の重複を削除しました。`;
    UI.showSuccessMessage(message);
    Logger.endProcess('重複チェック完了');
    
    // 処理後のデータ行数を取得
    const afterDataCount = listingSheet.getLastRow() - 1; // ヘッダー行を除く
    
    // 詳細な処理結果を表示
    UI.showResultMessage(
      '重複チェック完了',
      {
        removedCount: rowsToDelete.length,
        totalProcessed: dataRows.length
      },
      null
    );
    
    // 処理結果を返す
    return {
      success: true,
      message: message,
      stats: {
        removedCount: rowsToDelete.length,
        totalProcessed: dataRows.length
      }
    };
  } catch (error) {
    Logger.logError('重複チェック中にエラー: ' + error.message);
    UI.showErrorMessage('重複チェック中にエラーが発生しました: ' + error.message);
    return {
      success: false,
      message: '重複チェック中にエラーが発生しました: ' + error.message,
      stats: {
        removedCount: 0,
        totalProcessed: 0
      }
    };
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
 * @return {Object} 処理結果
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
    
    // Title列のインデックスを取得
    const titleColumnIndex = headerRow.indexOf('Title');
    if (titleColumnIndex === -1) {
      throw new Error('Title列が見つかりません。ヘッダー行に「Title」が含まれているか確認してください。');
    }
    
    // 設定を取得
    const settings = Config.getSettings();
    const characterLimit = settings.characterLimit;
    
    // 設定値が取得できない場合は処理を中止
    if (characterLimit === undefined || characterLimit === null) {
      throw new Error('文字数制限の設定値が取得できません。設定シートを確認してください。');
    }
    
    Logger.log(`文字数制限フィルター: ${characterLimit}文字以下を削除`);
    
    // 削除対象の行を特定
    let rowsToDelete = [];
    
    // フィルタリング処理
    dataRows.forEach((row, index) => {
      // 処理の進捗状況を更新（10%単位）
      if (index % Math.floor(Math.max(dataRows.length, 10) / 10) === 0) {
        UI.updateProgressBar(Math.floor((index / Math.max(dataRows.length, 1)) * 100));
      }
      
      const title = row[titleColumnIndex]; // Title列の値
      
      // 文字数チェック
      if (title && title.length <= characterLimit) {
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
    
    // 結果メッセージ
    const message = `文字数制限フィルタリングが完了しました。${rowsToDelete.length}件削除しました（${characterLimit}文字以上を削除）。`;
    UI.showSuccessMessage(message);
    Logger.endProcess('文字数制限フィルタリング完了');
    
    // 処理後のデータ行数を取得
    const afterDataCount = listingSheet.getLastRow() - 1; // ヘッダー行を除く
    
    // 詳細な処理結果を表示
    UI.showResultMessage(
      '文字数制限フィルタリング完了',
      {
        removedCount: rowsToDelete.length,
        characterLimit: characterLimit,
        totalProcessed: dataRows.length
      },
      `文字数制限: ${characterLimit}文字以上を削除`
    );
    
    // 処理結果を返す
    return {
      success: true,
      message: message,
      stats: {
        removedCount: rowsToDelete.length,
        characterLimit: characterLimit,
        totalProcessed: dataRows.length
      }
    };
  } catch (error) {
    Logger.logError('文字数制限フィルタリング中にエラー: ' + error.message);
    UI.showErrorMessage('文字数制限フィルタリング中にエラーが発生しました: ' + error.message);
    return {
      success: false,
      message: '文字数制限フィルタリング中にエラーが発生しました: ' + error.message,
      stats: {
        removedCount: 0,
        characterLimit: 0,
        totalProcessed: 0
      }
    };
  } finally {
    UI.hideProgressBar();
  }
};

/**
 * 所在地情報修正を実行する
 * @return {Object} 処理結果
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
    
    UI.updateProgressBar(100);
    
    // 結果メッセージ
    const message = `所在地情報の修正が完了しました。${updatedLocations.length}件の所在地情報を修正しました。`;
    UI.showSuccessMessage(message);
    Logger.endProcess('所在地情報修正完了');
    
    // 処理後のデータ行数を取得
    const afterDataCount = listingSheet.getLastRow() - 1; // ヘッダー行を除く
    
    // 詳細な処理結果を表示
    UI.showResultMessage(
      '所在地情報修正完了',
      {
        modifiedCount: updatedLocations.length,
        totalProcessed: dataRows.length
      },
      `適用パターン数: ${updatedLocations.length}件`
    );
    
    // 処理結果を返す
    return {
      success: true,
      message: message,
      stats: {
        modifiedCount: updatedLocations.length,
        totalProcessed: dataRows.length
      }
    };
  } catch (error) {
    Logger.logError('所在地情報修正中にエラー: ' + error.message);
    UI.showErrorMessage('所在地情報修正中にエラーが発生しました: ' + error.message);
    return {
      success: false,
      message: '所在地情報修正中にエラーが発生しました: ' + error.message,
      stats: {
        modifiedCount: 0,
        totalProcessed: 0
      }
    };
  } finally {
    UI.hideProgressBar();
  }
};

/**
 * 価格フィルタリングを実行する
 * @return {Object} 処理結果
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
    
    // StartPrice列のインデックスを取得
    const priceColumnIndex = headerRow.indexOf('StartPrice');
    if (priceColumnIndex === -1) {
      throw new Error('StartPrice列が見つかりません。ヘッダー行に「StartPrice」が含まれているか確認してください。');
    }
    
    // 設定を取得
    const settings = Config.getSettings();
    const priceThreshold = settings.priceThreshold;
    
    // 設定値が取得できない場合は処理を中止
    if (priceThreshold === undefined || priceThreshold === null) {
      throw new Error('価格下限の設定値が取得できません。設定シートを確認してください。');
    }
    
    Logger.log(`価格フィルター: $${priceThreshold}以下を削除`);
    
    // 削除対象の行を特定
    let rowsToDelete = [];
    
    // 価格フィルタリング処理
    dataRows.forEach((row, index) => {
      // 処理の進捗状況を更新（10%単位）
      if (index % Math.floor(Math.max(dataRows.length, 10) / 10) === 0) {
        UI.updateProgressBar(Math.floor((index / Math.max(dataRows.length, 1)) * 100));
      }
      
      let price = row[priceColumnIndex]; // StartPrice列の値
      
      // 文字列の場合は数値に変換
      if (typeof price === 'string') {
        // 数値以外の文字を削除（$や,など）
        price = parseFloat(price.replace(/[^0-9.]/g, ''));
      }
      
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
    
    // 結果メッセージ
    const message = `価格フィルタリングが完了しました。${rowsToDelete.length}件削除しました（${priceThreshold}ドル以下を削除）。`;
    UI.showSuccessMessage(message);
    Logger.endProcess('価格フィルタリング完了');
    
    // 処理後のデータ行数を取得
    const afterDataCount = listingSheet.getLastRow() - 1; // ヘッダー行を除く
    
    // 詳細な処理結果を表示
    UI.showResultMessage(
      '価格フィルタリング完了',
      {
        removedCount: rowsToDelete.length,
        priceThreshold: priceThreshold,
        totalProcessed: dataRows.length
      },
      `価格下限: ${priceThreshold}ドル以下を削除`
    );
    
    // 処理結果を返す
    return {
      success: true,
      message: message,
      stats: {
        removedCount: rowsToDelete.length,
        priceThreshold: priceThreshold,
        totalProcessed: dataRows.length
      }
    };
  } catch (error) {
    Logger.logError('価格フィルタリング中にエラー: ' + error.message);
    UI.showErrorMessage('価格フィルタリング中にエラーが発生しました: ' + error.message);
    return {
      success: false,
      message: '価格フィルタリング中にエラーが発生しました: ' + error.message,
      stats: {
        removedCount: 0,
        priceThreshold: 0,
        totalProcessed: 0
      }
    };
  } finally {
    UI.hideProgressBar();
  }
}; 