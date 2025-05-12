/**
 * eBay出品作業効率化ツール - UIモジュール
 * 
 * ユーザーインターフェース関連の処理を提供します。
 */

// UI名前空間
const UI = {};

/**
 * サイドバーを表示する
 */
UI.showSidebar = function() {
  const template = HtmlService.createTemplateFromFile('Sidebar');
  
  // テンプレートにバージョン情報を渡す
  template.version = APP_VERSION || 'v1.1.0';
  
  const html = template.evaluate()
    .setTitle(Config.UI.SIDEBAR_TITLE)
    .setWidth(Config.UI.SIDEBAR_WIDTH);
  
  SpreadsheetApp.getUi().showSidebar(html);
};

/**
 * CSVインポートダイアログを表示する
 */
UI.showImportDialog = function() {
  const html = HtmlService.createTemplateFromFile('ImportDialog')
    .evaluate()
    .setWidth(Config.UI.DIALOG_WIDTH)
    .setHeight(Config.UI.DIALOG_HEIGHT)
    .setTitle('CSVインポート');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'CSVインポート');
};

/**
 * CSVエクスポートダイアログを表示する
 */
UI.showExportDialog = function() {
  // 直接データインポートシートからCSVエクスポートを行う
  const html = HtmlService.createTemplateFromFile('ExportDialog')
    .evaluate()
    .setWidth(Config.UI.DIALOG_WIDTH)
    .setHeight(Config.UI.DIALOG_HEIGHT)
    .setTitle('CSVエクスポート');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'CSVエクスポート');
};

/**
 * 設定ダイアログを表示する
 */
UI.showSettingsDialog = function() {
  const html = HtmlService.createTemplateFromFile('SettingsDialog')
    .evaluate()
    .setWidth(Config.UI.DIALOG_WIDTH)
    .setHeight(Config.UI.DIALOG_HEIGHT)
    .setTitle('設定');
  
  SpreadsheetApp.getUi().showModalDialog(html, '設定');
};

/**
 * ヘルプダイアログを表示する
 */
UI.showHelpDialog = function() {
  const html = HtmlService.createTemplateFromFile('HelpDialog')
    .evaluate()
    .setWidth(Config.UI.DIALOG_WIDTH)
    .setHeight(Config.UI.DIALOG_HEIGHT)
    .setTitle('ヘルプ');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'ヘルプ');
};

/**
 * 成功メッセージを表示する
 * @param {string} message 表示するメッセージ
 * @param {string} url オプションのURL
 */
UI.showSuccessMessage = function(message, url) {
  try {
    // UIが利用可能かチェック
    let ui;
    try {
      ui = SpreadsheetApp.getUi();
    } catch (e) {
      // UIが利用できない場合はログのみに記録
      console.log('成功: ' + message);
      Logger.log('UI利用不可のためログのみに記録(成功): ' + message);
      return;
    }
    
    // サイドバーに通知を表示
    try {
      // サイドバー経由でメッセージを表示
      const script = `
        if (typeof showNotification === 'function') {
          showNotification('success', '${message}');
        } else {
          alert('${message}');
        }
        
        ${url ? `
          if (confirm('CSVをダウンロードしますか？')) {
            window.open('${url}', '_blank');
          }
        ` : ''}
      `;
      
      const html = HtmlService.createHtmlOutput(
        '<script>' + script + '</script>'
      );
      
      ui.showModelessDialog(html, '完了');
    } catch (e) {
      // サイドバーがない場合は通常のアラート
      if (url) {
        const response = ui.alert('成功', message + '\n\nCSVをダウンロードしますか？', ui.ButtonSet.YES_NO);
        if (response === ui.Button.YES) {
          // ダウンロードURLを開く
          const html = HtmlService.createHtmlOutput(
            '<script>window.open("' + url + '", "_blank"); google.script.host.close();</script>'
          )
          .setWidth(10)
          .setHeight(10);
          
          ui.showModalDialog(html, 'ダウンロード中...');
        }
      } else {
        ui.alert('成功', message, ui.ButtonSet.OK);
      }
    }
  } catch (finalError) {
    // 最終的なフォールバック - ログのみ
    console.log('成功メッセージ表示失敗: ' + message);
    Logger.log('成功メッセージ表示失敗: ' + message);
  }
};

/**
 * エラーメッセージを表示する
 * @param {string} message 表示するメッセージ
 */
UI.showErrorMessage = function(message) {
  try {
    // UIが利用可能かチェック
    let ui;
    try {
      ui = SpreadsheetApp.getUi();
    } catch (e) {
      // UIが利用できない場合はログのみに記録
      console.error('UI利用不可: ' + message);
      Logger.logError('UI利用不可のためログのみに記録: ' + message);
      return;
    }
    
    // サイドバーに通知を表示
    try {
      // サイドバー経由でメッセージを表示
      const script = `
        if (typeof showNotification === 'function') {
          showNotification('error', '${message}');
        } else {
          alert('エラー: ${message}');
        }
      `;
      
      const html = HtmlService.createHtmlOutput(
        '<script>' + script + '</script>'
      );
      
      ui.showModelessDialog(html, 'エラー');
    } catch (e) {
      // サイドバーがない場合は通常のアラート
      ui.alert('エラー', message, ui.ButtonSet.OK);
    }
  } catch (finalError) {
    // 最終的なフォールバック - ログのみ
    console.error('エラーメッセージ表示失敗: ' + message);
    Logger.logError('エラーメッセージ表示失敗: ' + message);
  }
};

/**
 * プログレスバーを表示する
 * @param {string} message 表示するメッセージ
 */
UI.showProgressBar = function(message) {
  try {
    // UIが利用可能かチェック
    try {
      SpreadsheetApp.getUi();
    } catch (e) {
      // UIが利用できない場合はログのみに記録
      console.log('処理中: ' + message);
      return;
    }
    
    // サイドバーを通じてプログレスバーを表示/更新
    try {
      // サイドバーのみに表示
      const html = HtmlService.createHtmlOutput(
        '<script>if (window.parent && window.parent.showProgress) { window.parent.showProgress("' + message + '", 0); }</script>'
      );
      SpreadsheetApp.getUi().showModelessDialog(html, '');
    } catch (e) {
      // エラーが発生した場合は無視（サイドバーが表示されていない可能性あり）
    }
  } catch (finalError) {
    // すべてのエラーを無視
  }
};

/**
 * プログレスバーを更新する
 * @param {number} percent 進捗パーセンテージ（0-100）
 */
UI.updateProgressBar = function(percent) {
  try {
    // UIが利用可能かチェック
    try {
      SpreadsheetApp.getUi();
    } catch (e) {
      // UIが利用できない場合は何もしない
      return;
    }
    
    // サイドバーのみに表示
    const html = HtmlService.createHtmlOutput(
      '<script>if (window.parent && window.parent.updateProgress) { window.parent.updateProgress(' + percent + '); }</script>'
    );
    SpreadsheetApp.getUi().showModelessDialog(html, '');
  } catch (e) {
    // エラーが発生した場合は無視
  }
};

/**
 * プログレスバーを非表示にする
 */
UI.hideProgressBar = function() {
  try {
    // UIが利用可能かチェック
    try {
      SpreadsheetApp.getUi();
    } catch (e) {
      // UIが利用できない場合は何もしない
      return;
    }
    
    // サイドバーのみに表示
    const html = HtmlService.createHtmlOutput(
      '<script>if (window.parent && window.parent.hideProgress) { window.parent.hideProgress(); }</script>'
    );
    SpreadsheetApp.getUi().showModelessDialog(html, '');
  } catch (e) {
    // エラーが発生した場合は無視
  }
};

/**
 * HTMLテンプレートに含めるJavaScriptを取得する
 */
UI.getJavaScript = function() {
  return HtmlService.createHtmlOutputFromFile('JavaScript').getContent();
};

/**
 * HTMLテンプレートに含めるCSSを取得する
 */
UI.getStylesheet = function() {
  return HtmlService.createHtmlOutputFromFile('Stylesheet').getContent();
}; 