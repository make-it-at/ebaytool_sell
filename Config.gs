/**
 * eBay出品作業効率化ツール - 設定モジュール
 * 
 * アプリケーション全体で使用する設定値を提供します。
 * 
 * バージョン: v1.3.7
 * 最終更新日: 2025-05-15
 */

// Config名前空間
const Config = {
  // シート名
  SHEET_NAMES: {
    IMPORT: 'データインポート',
    LISTING: '出品データ',  // 出品データシートを追加
    SETTINGS: '設定',
    LOG: 'ログ',
    TEST_DATA: 'テストデータ'  // テストデータ用シートの追加
  },
  
  // 各シートのヘッダー
  SHEET_HEADERS: {
    // データインポートシートのヘッダー
    IMPORT: [
      '商品名', 
      '価格($)', 
      '所在地', 
      'コンディション', 
      '出品者', 
      'URL',
      'リサーチ日',
      '処理結果'
    ],
    
    // エクスポート用eBayフォーマットのヘッダー
    EBAY_FORMAT: [
      'Action',
      'Item number',
      'Title',
      'Subtitle',
      'Category number',
      'Store category name 1',
      'Store category name 2',
      'Description',
      'Condition',
      'Picture URL',
      'Duration',
      'Format',
      'Start price',
      'Buy It Now price',
      'Quantity',
      'PayPal accepted',
      'PayPal email',
      'Immediate payment required',
      'Location',
      'Shipping service 1',
      'Shipping service cost 1',
      'Shipping service additional cost 1',
      'Shipping service 2',
      'Shipping service cost 2',
      'Shipping service additional cost 2',
      'Max dispatch time',
      'Returns accepted',
      'Refund',
      'Return shipping cost paid by'
    ]
  },
  
  // 設定のデフォルト値
  DEFAULT_SETTINGS: {
    CHARACTER_LIMIT: 20,   // 商品名文字数制限
    PRICE_THRESHOLD: 10,   // 価格下限（ドル）
    DUPLICATE_THRESHOLD: 80, // 重複判定の類似度閾値（%）
    NG_WORD_MODE: 'リスト全削除'  // NGワード処理モード
  },
  
  // テーブルのセクション開始行目安（実際の値は動的に計算）
  SECTION_START: {
    NG_WORDS: 2,         // NGワードテーブルの開始行（ヘッダー行の次）
    SETTINGS: 10         // 設定項目テーブルの開始行目安
  },
  
  // 最大処理行数
  MAX_ROWS: 10000,
  
  // UI関連設定
  UI: {
    SIDEBAR_TITLE: 'eBay出品ツール',
    SIDEBAR_WIDTH: 300,
    SIDEBAR_HEIGHT: 600,
    DIALOG_WIDTH: 600,
    DIALOG_HEIGHT: 400,
    THEME_COLOR: '#4CAF50' // グリーン基調
  }
};

/**
 * 設定値を取得する
 */
Config.getSettings = function() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(this.SHEET_NAMES.SETTINGS);
  
  // 設定シートがなければデフォルト値を返す
  if (!settingsSheet) {
    Logger.log('設定シートが見つかりません。デフォルト値を使用します。');
    return this.DEFAULT_SETTINGS;
  }
  
  try {
    // 設定シートのヘッダー列が存在するか確認
    const headerRow = settingsSheet.getRange(1, 1, 1, 5).getValues()[0];
    if (headerRow.length < 5 || !headerRow[0] || !headerRow[2] || !headerRow[3] || !headerRow[4]) {
      throw new Error('設定シートのヘッダー行が正しくありません。「リスト削除」「削除ワード」「文字数制限」「価格下限」「重複類似度閾値」の列が必要です');
    }
    
    // リスト削除NGワードを取得（A列）
    let deleteListNgWords = [];
    let row = 2; // ヘッダー行の次から開始
    while (row <= 100) {
      const value = settingsSheet.getRange(row, 1).getValue();
      if (!value || value === '') {
        break; // 空の行で終了
      }
      deleteListNgWords.push(value);
      row++;
    }
    
    // 削除ワードを取得（B列）
    let deletePartNgWords = [];
    row = 2; // ヘッダー行の次から開始
    while (row <= 100) {
      const value = settingsSheet.getRange(row, 2).getValue();
      if (!value || value === '') {
        break; // 空の行で終了
      }
      deletePartNgWords.push(value);
      row++;
    }
    
    // 文字数制限値を取得（C列2行目）
    let characterLimit = this.DEFAULT_SETTINGS.CHARACTER_LIMIT; // デフォルト値をあらかじめセット
    try {
      let characterLimitRaw = settingsSheet.getRange(2, 3).getValue();
      let parsedCharacterLimit = parseInt(characterLimitRaw);
      if (!isNaN(parsedCharacterLimit) && parsedCharacterLimit > 0) {
        characterLimit = parsedCharacterLimit;
      } else {
        Logger.log(`文字数制限が不正です: ${characterLimitRaw}。デフォルト値(${this.DEFAULT_SETTINGS.CHARACTER_LIMIT})を使用します。`);
      }
    } catch (e) {
      Logger.logError(`文字数制限の取得中にエラー: ${e.message}。デフォルト値(${this.DEFAULT_SETTINGS.CHARACTER_LIMIT})を使用します。`);
    }
    
    // 価格下限値を取得（D列2行目）
    let priceThreshold = this.DEFAULT_SETTINGS.PRICE_THRESHOLD; // デフォルト値をあらかじめセット
    try {
      let priceThresholdRaw = settingsSheet.getRange(2, 4).getValue();
      let parsedPriceThreshold = parseFloat(priceThresholdRaw);
      if (!isNaN(parsedPriceThreshold) && parsedPriceThreshold >= 0) {
        priceThreshold = parsedPriceThreshold;
      } else {
        Logger.log(`価格下限が不正です: ${priceThresholdRaw}。デフォルト値(${this.DEFAULT_SETTINGS.PRICE_THRESHOLD})を使用します。`);
      }
    } catch (e) {
      Logger.logError(`価格下限の取得中にエラー: ${e.message}。デフォルト値(${this.DEFAULT_SETTINGS.PRICE_THRESHOLD})を使用します。`);
    }
    
    // 重複閾値を取得（E列2行目）
    let duplicateThreshold = this.DEFAULT_SETTINGS.DUPLICATE_THRESHOLD; // デフォルト値をあらかじめセット
    try {
      let duplicateThresholdRaw = settingsSheet.getRange(2, 5).getValue();
      let parsedDuplicateThreshold = parseInt(duplicateThresholdRaw);
      if (!isNaN(parsedDuplicateThreshold) && parsedDuplicateThreshold > 0 && parsedDuplicateThreshold <= 100) {
        duplicateThreshold = parsedDuplicateThreshold;
      } else {
        Logger.log(`重複閾値が不正です: ${duplicateThresholdRaw}。デフォルト値(${this.DEFAULT_SETTINGS.DUPLICATE_THRESHOLD})を使用します。`);
      }
    } catch (e) {
      Logger.logError(`重複閾値の取得中にエラー: ${e.message}。デフォルト値(${this.DEFAULT_SETTINGS.DUPLICATE_THRESHOLD})を使用します。`);
    }
    
    // 設定値をログ出力（デバッグ用）
    Logger.log(`設定を読み込みました: リスト削除NGワード=${deleteListNgWords.length}件, 部分削除NGワード=${deletePartNgWords.length}件, 文字数制限=${characterLimit}, 価格下限=${priceThreshold}, 重複閾値=${duplicateThreshold}`);
    
    return {
      deleteListNgWords: deleteListNgWords,
      deletePartNgWords: deletePartNgWords,
      characterLimit: characterLimit,
      priceThreshold: priceThreshold,
      duplicateThreshold: duplicateThreshold
    };
  } catch (error) {
    // エラーが発生した場合はデフォルト値を返す
    Logger.logError('設定値の取得中にエラーが発生しました: ' + error.message + '。デフォルト値を使用します。');
    return {
      deleteListNgWords: [],
      deletePartNgWords: [],
      characterLimit: this.DEFAULT_SETTINGS.CHARACTER_LIMIT,
      priceThreshold: this.DEFAULT_SETTINGS.PRICE_THRESHOLD,
      duplicateThreshold: this.DEFAULT_SETTINGS.DUPLICATE_THRESHOLD
    };
  }
};

/**
 * 設定値を保存する
 */
Config.saveSettings = function(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(this.SHEET_NAMES.SETTINGS);
  
  if (!settingsSheet) {
    return false;
  }
  
  // NGワードリストの保存（提供された場合のみ）
  if (settings.ngWords && Array.isArray(settings.ngWords)) {
    // まず既存のNGワードをクリア
    let row = this.SECTION_START.NG_WORDS;
    while (row <= 100) {
      const value = settingsSheet.getRange(row, 1).getValue();
      if (!value || value === '' || value === '設定項目') {
        break; // 空の行または次のセクションヘッダーで終了
      }
      settingsSheet.getRange(row, 1, 1, 2).clearContent();
      row++;
    }
    
    // 新しいNGワードを設定
    if (settings.ngWords.length > 0) {
      const newNgWords = settings.ngWords.map(word => [word, '']);
      settingsSheet.getRange(this.SECTION_START.NG_WORDS, 1, newNgWords.length, 2).setValues(newNgWords);
    }
  }
  
  // 設定項目テーブルの位置を検索
  let settingsRow = this.SECTION_START.SETTINGS;
  let found = false;
  
  while (settingsRow <= 100 && !found) {
    const value = settingsSheet.getRange(settingsRow, 1).getValue();
    if (value === '設定項目') {
      found = true;
      break;
    }
    settingsRow++;
  }
  
  // 設定項目が見つかった場合、値を更新
  if (found) {
    const valueRow = settingsRow + 1;
    
    // 提供された場合のみ値を更新
    if (settings.ngWordMode) {
      settingsSheet.getRange(valueRow, 2).setValue(settings.ngWordMode);
    }
    
    if (settings.characterLimit) {
      settingsSheet.getRange(valueRow, 3).setValue(settings.characterLimit);
    }
    
    if (settings.priceThreshold) {
      settingsSheet.getRange(valueRow, 4).setValue(settings.priceThreshold);
    }
    
    if (settings.duplicateThreshold) {
      settingsSheet.getRange(valueRow, 5).setValue(settings.duplicateThreshold);
    }
  }
  
  return true;
}; 