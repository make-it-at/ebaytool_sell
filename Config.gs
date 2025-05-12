/**
 * eBay出品作業効率化ツール - 設定モジュール
 * 
 * アプリケーション全体で使用する設定値を提供します。
 * 
 * バージョン: v1.2.0
 * 最終更新日: 2024-07-16
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
  
  // 設定がなければデフォルト値を返す
  if (!settingsSheet) {
    return this.DEFAULT_SETTINGS;
  }
  
  // NGワードリストを取得（A列の値をすべて取得）
  let ngWords = [];
  let row = this.SECTION_START.NG_WORDS;
  while (row <= 100) { // 最大100行まで検索
    const value = settingsSheet.getRange(row, 1).getValue();
    if (!value || value === '' || value === '設定項目') {
      break; // 空の行または次のセクションヘッダーで終了
    }
    ngWords.push(value);
    row++;
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
  
  // 設定項目が見つからない場合はデフォルト値を返す
  if (!found) {
    return {
      ngWords: ngWords.filter(word => word.trim() !== ''),
      ...this.DEFAULT_SETTINGS
    };
  }
  
  // 設定値は見出しの1行下から取得
  const valueRow = settingsRow + 1;
  const ngWordMode = settingsSheet.getRange(valueRow, 2).getValue() || this.DEFAULT_SETTINGS.NG_WORD_MODE;
  const characterLimit = parseInt(settingsSheet.getRange(valueRow, 3).getValue()) || this.DEFAULT_SETTINGS.CHARACTER_LIMIT;
  const priceThreshold = parseFloat(settingsSheet.getRange(valueRow, 4).getValue()) || this.DEFAULT_SETTINGS.PRICE_THRESHOLD;
  const duplicateThreshold = parseInt(settingsSheet.getRange(valueRow, 5).getValue()) || this.DEFAULT_SETTINGS.DUPLICATE_THRESHOLD;
  
  return {
    ngWords: ngWords.filter(word => word.trim() !== ''),
    characterLimit: characterLimit,
    priceThreshold: priceThreshold,
    duplicateThreshold: duplicateThreshold,
    ngWordMode: ngWordMode
  };
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