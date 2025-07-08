/**
 * eBay出品作業効率化ツール - UIモジュール
 * 
 * ユーザーインターフェース関連の処理を提供します。
 * 
 * バージョン: v1.5.13
 * 最終更新日: 2025-06-14
 * 更新内容: 処理完了メッセージから詳細情報を削除
 */

// UI名前空間
const UI = {};

// 直近の成功メッセージを保存するグローバル変数
let LAST_RESULT_MESSAGE = null;

// プログレスバーの状態を保持するグローバル変数
let _progressState = {
  isVisible: false,
  message: '',
  percent: 0,
  completion: false
};

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
 * プログレスバーの状態をリセットする
 * 処理開始前に既存のプログレスバーの状態をリセットするために使用
 */
UI.resetProgressState = function() {
  try {
    // グローバル変数として保存されるプログレス状態をリセット
    _progressState = {
      isVisible: true,
      message: '処理を開始しています...',
      percent: 0,
      completion: false  // 重要: 完了フラグを必ずfalseに設定
    };
    
    // サイドバーHTMLへのスクリプト実行は不要
    // クライアント側のgetProgressState()が定期的に状態を取得するため
    
    Logger.log('プログレスバーをリセットしました');
  } catch (e) {
    console.error('Error in resetProgressState:', e);
    Logger.logError('プログレスバーリセット失敗: ' + e.message);
  }
  
  return true;
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
    
    // ログに記録
    Logger.log('成功: ' + message);
    
    // サイドバー表示用の処理は不要
    // getLastResultMessage()がクライアント側から定期的に呼ばれるため
  } catch (error) {
    // ログのみに記録
    Logger.log('成功メッセージ表示失敗: ' + error.message);
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
    // 直近のメッセージを保存
    LAST_RESULT_MESSAGE = {
      type: 'error',
      title: 'エラー',
      content: message
    };
    
    // ログに記録
    Logger.logError('エラー: ' + message);
    
    // サイドバー表示用の処理は不要
    // getLastResultMessage()がクライアント側から定期的に呼ばれるため
  } catch (error) {
    // ログのみに記録
    Logger.logError('エラーメッセージ表示失敗: ' + error.message);
  }
};

/**
 * プログレス状態をPropertiesServiceに保存する
 */
function saveProgressState(state) {
  PropertiesService.getScriptProperties().setProperty('PROGRESS_STATE', JSON.stringify(state));
}

/**
 * プログレス状態をPropertiesServiceから取得する
 */
function loadProgressState() {
  const str = PropertiesService.getScriptProperties().getProperty('PROGRESS_STATE');
  if (!str) return null;
  try {
    return JSON.parse(str);
  } catch (e) {
    return null;
  }
}

/**
 * プログレスバーを表示する
 * @param {string} message 表示するメッセージ
 * @param {boolean} reset リセットフラグ（省略可）
 */
UI.showProgressBar = function(message, reset) {
  try {
    // サイドバーに通知を表示（可能な場合）
    try {
      // エスケープ処理を行う
      const escapedMessage = message.replace(/'/g, "\\'").replace(/\n/g, "\\n");
      
      // サイドバーに表示するスクリプト
      const script = `
        if (typeof showProgressBar === 'function') {
          showProgressBar('${escapedMessage}', ${reset ? 'true' : 'false'});
        }
      `;
      
      // スクリプトを実行
      ScriptApp.run(() => {
        try {
          const html = HtmlService.createHtmlOutput(`<script>${script}</script>`);
        } catch (e) {
          // エラーは無視
        }
      });
    } catch (e) {
      // エラーは無視
    }
  } catch (error) {
    // ログのみに記録
    Logger.log('プログレスバー表示エラー: ' + error.message);
  }
};

/**
 * プログレスバーを更新する
 * @param {number} progress 進捗率（0-100）
 * @param {boolean} completed 完了フラグ（省略可）- v1.5.6で廃止
 */
UI.updateProgressBar = function(progress, completed) {
  try {
    // 新しいプログレス状態を設定
    _progressState = {
      isVisible: true,
      message: _progressState ? _progressState.message : '処理中...',
      percent: progress,
      completion: false  // 常にfalseに設定（チェックマークを表示しない）
    };
    
    // スピナーが消えないように明示的に表示状態を維持するスクリプトを追加
    try {
      // エスケープ処理を行う
      const script = `
        const spinner = document.querySelector('.spinner');
        if (spinner) {
          spinner.style.display = 'block';
          // スタイルの再計算を強制
          void spinner.offsetHeight;
        }
      `;
      
      // スクリプトを実行
      ScriptApp.run(() => {
        try {
          const html = HtmlService.createHtmlOutput(`<script>${script}</script>`);
        } catch (e) {
          // エラーは無視
        }
      });
    } catch (e) {
      // エラーは無視
    }
    
    // サイドバーは自動的に状態を取得するため追加処理は不要
  } catch (error) {
    // ログのみに記録
    Logger.log('プログレスバー更新エラー: ' + error.message);
  }
};

/**
 * プログレスバーを非表示にする
 */
UI.hideProgressBar = function() {
  try {
    // サイドバーに通知を表示（可能な場合）
    try {
      // サイドバーに表示するスクリプト
      const script = `
        if (typeof hideProgressBar === 'function') {
          hideProgressBar();
        }
      `;
      
      // スクリプトを実行
      ScriptApp.run(() => {
        try {
          const html = HtmlService.createHtmlOutput(`<script>${script}</script>`);
        } catch (e) {
          // エラーは無視
        }
      });
    } catch (e) {
      // エラーは無視
    }
  } catch (error) {
    // ログのみに記録
    Logger.log('プログレスバー非表示エラー: ' + error.message);
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
 * 処理結果メッセージを表示する
 * フィルタリング処理などの結果をサイドバーに表示するための関数
 * @param {string} title メッセージのタイトル
 * @param {Object} stats 処理の統計情報（削除件数、修正件数など）
 * @param {string} additionalInfo 追加情報（オプション）
 */
UI.showResultMessage = function(title, stats, additionalInfo) {
  try {
    // 直近のメッセージを保存 - これをサイドバーがgetLastResultMessageで取得する
    LAST_RESULT_MESSAGE = {
      type: 'result',
      title: title,
      stats: stats || { removedCount: 0, modifiedCount: 0 },
      additionalInfo: additionalInfo || ''
    };
    
    // ログに記録
    Logger.log(`処理結果: ${title} (削除: ${stats?.removedCount || 0}件, 修正: ${stats?.modifiedCount || 0}件)`);
    
    // トーストメッセージは使用しない
  } catch (error) {
    // エラーはログのみに記録
    Logger.logError(`結果メッセージ表示失敗: ${error.message}`);
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
  return _progressState;
}

/**
 * プログレストラッカー（複数ステージの進捗管理）
 */
UI.ProgressTracker = {
  /**
   * プログレストラッカーを初期化する
   * @param {Object} stages ステージごとの重み付け（ステージ名: 重み）
   * @return {Object} プログレストラッカーインスタンス
   */
  init: function(stages) {
    const totalWeight = Object.values(stages).reduce((sum, weight) => sum + weight, 0);
    const stageWeights = {};
    
    // 各ステージの相対的な重みを計算
    for (const stageName in stages) {
      stageWeights[stageName] = stages[stageName] / totalWeight;
    }
    
    // 累積進捗率を計算（ステージの開始地点の進捗率）
    const stageCumulativeProgress = {};
    let cumulativeProgress = 0;
    
    for (const stageName in stageWeights) {
      stageCumulativeProgress[stageName] = cumulativeProgress;
      cumulativeProgress += stageWeights[stageName] * 100;
    }
    
    return {
      stageWeights: stageWeights,
      stageCumulativeProgress: stageCumulativeProgress,
      currentStage: null,
      currentStageStartProgress: 0,
      currentStageWeight: 0,
      
      /**
       * ステージを開始する
       * @param {string} stageName ステージ名
       * @param {string} message 表示するメッセージ
       */
      startStage: function(stageName, message) {
        this.currentStage = stageName;
        this.currentStageStartProgress = this.stageCumulativeProgress[stageName];
        this.currentStageWeight = this.stageWeights[stageName] * 100;
        
        // プログレスバーを更新
        UI.updateProgressBar(this.currentStageStartProgress);
        
        // メッセージを表示（指定がある場合）
        if (message) {
          UI.showProgressBar(message);
        }
      },
      
      /**
       * 現在のステージの進捗を更新する
       * @param {number} stageProgress ステージ内の進捗率（0-100）
       */
      updateStageProgress: function(stageProgress) {
        if (this.currentStage === null) return;
        
        // ステージ内の進捗を全体の進捗に変換
        const overallProgress = this.currentStageStartProgress + 
                               (stageProgress / 100) * this.currentStageWeight;
        
        // プログレスバーを更新
        UI.updateProgressBar(Math.min(overallProgress, 99.9)); // 100%は完了時のみ
      },
      
      /**
       * 現在のステージを完了する
       */
      completeStage: function() {
        if (this.currentStage === null) return;
        
        // ステージの完了地点の進捗率を計算
        const stageEndProgress = this.currentStageStartProgress + this.currentStageWeight;
        
        // プログレスバーを更新
        UI.updateProgressBar(stageEndProgress);
      },
      
      /**
       * 全ての処理を完了する
       * @param {string} message 完了メッセージ（オプション）
       */
      complete: function(message) {
        // 進捗率を100%に設定
        UI.updateProgressBar(100, true);
        
        // メッセージを表示（指定がある場合）
        if (message) {
          UI.showProgressBar(message);
        }
      }
    };
  }
}; 