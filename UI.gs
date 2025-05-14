/**
 * eBay出品作業効率化ツール - UIモジュール
 * 
 * ユーザーインターフェース関連の処理を提供します。
 * 
 * バージョン: v1.3.1
 * 最終更新日: 2025-05-14
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
 * テストデータ設定についてのヘルプダイアログを表示
 */
UI.showTestDataHelpDialog = function() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 15px;">
      <h2>テストデータの設定方法</h2>
      <p>テストデータを「テストデータ」シートに設定するには、次の方法があります：</p>
      
      <ol>
        <li>ローカルのCSVファイル（<code>/Users/kyosukemakita/Documents/Cursor/ebaytool_sell/eBay Export May 12 2025.csv</code>）を開く</li>
        <li>そのデータをコピーして、スプレッドシートの「テストデータ」シートに貼り付ける</li>
        <li>データはヘッダー行を含み、6行目以降に貼り付けてください（1-5行目は情報表示用）</li>
      </ol>
      
      <p>設定後、サイドバーの「テスト用CSVインポート」ボタンをクリックすると、
      「テストデータ」シートのデータが「出品データ」シートに転記されます。</p>
      
      <div style="margin-top: 20px; color: #666;">
        <p><strong>注意：</strong> Google Apps Scriptではローカルファイルシステムに直接アクセスできないため、
        この方法でテストデータを設定しています。</p>
      </div>
      
      <div style="margin-top: 20px; text-align: center;">
        <button onclick="google.script.host.close();" 
                style="padding: 8px 15px; background: #4285f4; color: white; 
                       border: none; border-radius: 3px; cursor: pointer;">
          閉じる
        </button>
      </div>
    </div>
  `)
  .setWidth(500)
  .setHeight(400)
  .setTitle('テストデータ設定ヘルプ');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'テストデータ設定ヘルプ');
};

/**
 * 成功メッセージを表示する
 * @param {string} message 表示するメッセージ
 */
UI.showSuccessMessage = function(message) {
  try {
    // UIが利用可能かチェック
    try {
      SpreadsheetApp.getUi();
    } catch (e) {
      // UIが利用できない場合はログのみに記録
      console.log('成功: ' + message);
      Logger.log('UI利用不可のためログのみに記録(成功): ' + message);
      return;
    }
    
    // サイドバーに通知を表示
    try {
      // サイドバーが存在するかチェック
      const sidebar = HtmlService.createHtmlOutput(
        '<script>function checkSidebar() { return (typeof showNotification === "function"); }</script>'
      );
      
      // サイドバー経由でメッセージを表示するコールバック関数
      const callback = function(hasSidebar) {
        if (hasSidebar) {
          const script = `showNotification('success', '${message.replace(/'/g, "\\'")}');`;
          ScriptApp.run(() => {
            try {
              const html = HtmlService.createHtmlOutput(`<script>${script}</script>`);
              // 非表示のカスタム関数実行（ダイアログなし）
              SpreadsheetApp.getActive().toast(message, '成功', 3);
            } catch (e) {
              Logger.log('サイドバー通知表示エラー: ' + e.message);
            }
          });
        } else {
          // サイドバーがない場合はtoastのみ使用
          SpreadsheetApp.getActive().toast(message, '成功', 3);
        }
      };
      
      // トーストのみを使用
      SpreadsheetApp.getActive().toast(message, '成功', 3);
    } catch (e) {
      // トーストのみを使用
      SpreadsheetApp.getActive().toast(message, '成功', 3);
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
    try {
      SpreadsheetApp.getUi();
    } catch (e) {
      // UIが利用できない場合はログのみに記録
      console.error('UI利用不可: ' + message);
      Logger.logError('UI利用不可のためログのみに記録: ' + message);
      return;
    }
    
    // トーストを使用してメッセージを表示（ダイアログなし）
    SpreadsheetApp.getActive().toast(message, 'エラー', 5);
    
    // サイドバーに通知を表示（可能な場合）
    try {
      const script = `
        if (typeof showNotification === 'function') {
          showNotification('error', '${message.replace(/'/g, "\\'")}');
        }
      `;
      
      // サイドバーへのメッセージ送信を試行（エラーは無視）
      ScriptApp.run(() => {
        try {
          const html = HtmlService.createHtmlOutput(`<script>${script}</script>`);
        } catch (e) {
          // サイドバー通知に失敗してもトーストには表示されているので無視
        }
      });
    } catch (e) {
      // サイドバー通知に失敗してもトーストには表示されているので無視
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