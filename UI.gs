/**
 * eBay出品作業効率化ツール - UIモジュール
 * 
 * ユーザーインターフェース関連の処理を提供します。
 * 
 * バージョン: v1.3.4
 * 最終更新日: 2025-05-28
 */

// UI名前空間
const UI = {};

// 直近の成功メッセージを保存するグローバル変数
let LAST_RESULT_MESSAGE = null;

// 直近のプログレス状態を保存するグローバル変数
let LAST_PROGRESS_STATE = null;

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
    // 直近のメッセージを保存
    LAST_RESULT_MESSAGE = {
      type: 'success',
      title: '処理完了',
      content: message
    };
    // UIが利用可能かチェック
    try {
      SpreadsheetApp.getUi();
    } catch (e) {
      // UIが利用できない場合はログのみに記録
      console.log('成功: ' + message);
      Logger.log('UI利用不可のためログのみに記録(成功): ' + message);
      return;
    }
    
    // トースト通知は表示しない（削除）
    // ログのみに記録
    Logger.log('成功: ' + message);
  } catch (finalError) {
    // 最終的なフォールバック - ログのみ
    console.log('成功メッセージ表示失敗: ' + message);
    Logger.log('成功メッセージ表示失敗: ' + message);
  }
};

/**
 * 詳細な成功メッセージを表示する
 * タイトルと詳細内容の両方を表示し、サイドバーに結果メッセージとして表示される
 * @param {string} title メッセージのタイトル
 * @param {string} details 詳細内容
 */
UI.showDetailedSuccessMessage = function(title, details) {
  try {
    // UIが利用可能かチェック
    try {
      SpreadsheetApp.getUi();
    } catch (e) {
      // UIが利用できない場合はログのみに記録
      console.log(`成功: ${title}\n${details}`);
      Logger.log(`UI利用不可のためログのみに記録(成功): ${title}\n${details}`);
      return;
    }
    
    // サイドバーに通知を表示
    try {
      // エスケープ処理を行う
      const escapedTitle = title.replace(/'/g, "\\'").replace(/\n/g, "\\n");
      const escapedDetails = details.replace(/'/g, "\\'").replace(/\n/g, "\\n");
      
      // サイドバーとして表示するスクリプト
      const script = `
        if (typeof showResultMessage === 'function') {
          showResultMessage('success', '${escapedTitle}', '${escapedDetails}');
        } else if (typeof showNotification === 'function') {
          showNotification('success', '${escapedTitle}', false);
        }
      `;
      
      // スクリプトを実行
      ScriptApp.run(() => {
        try {
          const html = HtmlService.createHtmlOutput(`<script>${script}</script>`);
          // 非表示のカスタム関数実行（ダイアログなし）
          SpreadsheetApp.getActive().toast(title, '成功', 3);
        } catch (e) {
          Logger.log('サイドバー通知表示エラー: ' + e.message);
          // トーストのみを使用
          SpreadsheetApp.getActive().toast(title, '成功', 3);
        }
      });
    } catch (e) {
      // トーストのみを使用
      SpreadsheetApp.getActive().toast(title, '成功', 3);
    }
  } catch (finalError) {
    // 最終的なフォールバック - ログのみ
    console.log(`成功メッセージ表示失敗: ${title}\n${details}`);
    Logger.log(`成功メッセージ表示失敗: ${title}\n${details}`);
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
    
    // 現在のプログレスバー状態を保存
    LAST_PROGRESS_STATE = {
      message: message,
      percent: 0,
      isVisible: true
    };
    
    // ログに記録
    Logger.log(`プログレスバー表示: ${message}`);
    
    try {
      // サイドバー直接更新
      const script = `
        if (typeof showProgress === 'function') {
          showProgress('${message.replace(/'/g, "\\'")}', 0);
        }
      `;
      
      // サイドバーにスクリプトを送信（モードレスダイアログを使用）
      const html = HtmlService.createHtmlOutput(`<script>${script}</script>`)
        .setWidth(1).setHeight(1);
      SpreadsheetApp.getUi().showModelessDialog(html, '');
    } catch (e) {
      Logger.log('プログレスバー表示エラー: ' + e.message);
    }
  } catch (finalError) {
    // すべてのエラーを無視
    Logger.log('プログレスバー表示で致命的エラー: ' + finalError.message);
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
    
    // 現在のプログレスバー状態を更新
    if (LAST_PROGRESS_STATE) {
      LAST_PROGRESS_STATE.percent = percent;
    }
    
    // ログに進捗を記録（10%単位）
    if (percent % 10 === 0) {
      Logger.log(`プログレス更新: ${percent}%`);
    }
    
    // サイドバーのshowProgressMessageを呼び出す
    const functionName = "updateSidebarProgressBar";
    const args = [percent];
    
    // 直接サイドバー関数を呼び出す（グローバル関数経由）
    updateSidebarProgressBar(percent);
  } catch (e) {
    // エラーが発生した場合は無視
    Logger.log('プログレスバー更新エラー: ' + e.message);
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
    
    // プログレスバー状態をリセット
    if (LAST_PROGRESS_STATE) {
      LAST_PROGRESS_STATE.isVisible = false;
    }
    
    // ログに記録
    Logger.log('プログレスバー非表示');
    
    // 直接サイドバー関数を呼び出す（グローバル関数経由）
    hideSidebarProgressBar();
  } catch (e) {
    // エラーが発生した場合は無視
    Logger.log('プログレスバー非表示エラー: ' + e.message);
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

/**
 * 処理結果メッセージをサイドバーに表示する
 * フィルタリング処理の詳細な結果を表示し、サイドバーに固定表示する
 * @param {string} title メッセージのタイトル
 * @param {Object} stats 統計情報（削除件数、修正件数など）
 * @param {string} [additionalInfo] 追加情報（オプション）
 */
UI.showResultMessage = function(title, stats, additionalInfo) {
  try {
    // 統計情報のチェックと初期値設定
    stats = stats || {};
    const removedCount = stats.removedCount || 0;
    const modifiedCount = stats.modifiedCount || 0;
    const totalProcessed = stats.totalProcessed || 0;
    const beforeCount = stats.beforeCount || 0;
    const afterCount = stats.afterCount || 0;
    
    // フォーマットされた結果メッセージを作成
    let resultMessage = `<div class="result-message">`;
    resultMessage += `<h3>${title}</h3>`;
    
    // 処理前後のデータ数が提供されている場合は表示
    if (beforeCount > 0 || afterCount > 0) {
      resultMessage += `<p>データ数: ${beforeCount}件 → ${afterCount}件</p>`;
    } else {
      resultMessage += `<p>処理件数: ${totalProcessed}件</p>`;
    }
    
    if (removedCount > 0) {
      resultMessage += `<p>削除件数: ${removedCount}件</p>`;
    }
    
    if (modifiedCount > 0) {
      resultMessage += `<p>修正件数: ${modifiedCount}件</p>`;
    }
    
    // 追加情報（設定値など）がある場合は表示
    if (additionalInfo) {
      resultMessage += `<p class="additional-info">${additionalInfo}</p>`;
    }
    
    resultMessage += `</div>`;
    
    // エスケープ処理を行う
    const escapedResultMessage = resultMessage.replace(/'/g, "\\'").replace(/\n/g, "\\n");
    
    // サイドバーに通知を表示
    const script = `showResultMessageInSidebar('${escapedResultMessage}');`;
    try {
      const html = HtmlService.createHtmlOutput(`<script>${script}</script>`);
      // トースト通知は表示しない（削除）
      
      // ログにも記録
      Logger.log(`処理結果表示: ${title} (削除: ${removedCount}件, 修正: ${modifiedCount}件)`);
    } catch (e) {
      Logger.log('サイドバー結果表示エラー: ' + e.message);
      // トースト通知は表示しない（削除）
    }
  } catch (finalError) {
    // 最終的なフォールバック - ログのみ
    console.log('結果メッセージ表示失敗: ' + title);
    Logger.log('結果メッセージ表示失敗: ' + title);
  }
};

/**
 * サイドバーから取得するための直近のメッセージ取得関数
 */
function getLastResultMessage() {
  return LAST_RESULT_MESSAGE;
}

/**
 * サイドバーからのプログレスバー状態取得用グローバル関数
 */
function getProgressState() {
  return LAST_PROGRESS_STATE;
}

/**
 * サイドバーのプログレスバーを更新するグローバル関数
 */
function updateSidebarProgressBar(percent) {
  try {
    const script = `
      if (typeof updateProgress === 'function') {
        updateProgress(${percent});
      }
    `;
    const output = HtmlService.createHtmlOutput(`<script>${script}</script>`);
    SpreadsheetApp.getUi().showModelessDialog(output, "進捗を更新中...");
    return true;
  } catch (e) {
    Logger.log('サイドバープログレス更新エラー: ' + e.message);
    return false;
  }
}

/**
 * サイドバーのプログレスバーを非表示にするグローバル関数
 */
function hideSidebarProgressBar() {
  try {
    const script = `
      if (typeof hideProgress === 'function') {
        hideProgress();
      }
    `;
    const output = HtmlService.createHtmlOutput(`<script>${script}</script>`);
    SpreadsheetApp.getUi().showModelessDialog(output, "プログレスを閉じています...");
    return true;
  } catch (e) {
    Logger.log('サイドバープログレス非表示エラー: ' + e.message);
    return false;
  }
} 