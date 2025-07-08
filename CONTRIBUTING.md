# コントリビューションガイド

eBay出品管理ツールの開発にご協力いただき、ありがとうございます。

## 開発環境のセットアップ

### 必要なツール

- Node.js (v14以上)
- Google Apps Script CLI (`@google/clasp`)
- Git

### セットアップ手順

1. リポジトリをクローン
```bash
git clone https://github.com/make-it-at/ebaytool_sell.git
cd ebaytool_sell
```

2. 依存関係をインストール
```bash
npm install
```

3. Google Apps Script CLIの設定
```bash
clasp login
```

4. プロジェクトの設定
```bash
# 既存のGASプロジェクトにリンク
clasp clone [SCRIPT_ID]
```

## 開発フロー

### ブランチ戦略

- `main`: 本番環境用（安定版）
- `develop`: 開発用（最新の機能）
- `feature/*`: 新機能開発用
- `hotfix/*`: 緊急修正用

### 開発手順

1. developブランチから新しいブランチを作成
```bash
git checkout develop
git pull origin develop
git checkout -b feature/new-feature
```

2. 開発・テスト
```bash
# コードの変更
# テストの実行
npm run push  # GASにプッシュ
```

3. コミット
```bash
git add .
git commit -m "feat: 新機能の説明"
```

4. プルリクエストの作成
```bash
git push origin feature/new-feature
# GitHubでプルリクエストを作成
```

## コーディング規約

### ファイル構成

- `App.gs`: メイン処理とグローバル変数
- `Config.gs`: 設定値とNGワード定義
- `Filters.gs`: フィルタリング処理
- `ImportExport.gs`: CSV入出力処理
- `UI.gs`: ユーザーインターフェース処理
- `Logger.gs`: ログ処理
- `Sidebar.html`: サイドバーUI

### 命名規則

- 関数名: キャメルケース（`processData`）
- 変数名: キャメルケース（`userData`）
- 定数: アッパーケース（`MAX_ITEMS`）
- グローバル変数: アッパーケース（`ALL_PROCESSES_EXECUTED`）

### コメント

- 関数の説明は必須
- 複雑な処理には詳細なコメントを追加
- 日本語でのコメントを推奨

```javascript
/**
 * CSVデータをフィルタリングする
 * @param {Array} data - 処理対象のデータ配列
 * @param {Object} options - フィルタリングオプション
 * @return {Array} フィルタリング後のデータ配列
 */
function filterData(data, options) {
  // 処理の詳細...
}
```

## バージョン管理

### バージョン番号

セマンティックバージョニングを使用：`MAJOR.MINOR.PATCH`

- MAJOR: 破壊的変更
- MINOR: 新機能追加
- PATCH: バグ修正

### リリース手順

1. バージョン番号の更新
   - `package.json`
   - `App.gs` の `APP_VERSION`
   - `README.md`

2. CHANGELOGの更新

3. タグの作成
```bash
git tag v1.5.14
git push origin v1.5.14
```

## テスト

### 手動テスト

1. CSVインポート機能
2. 各フィルタリング機能
3. CSVエクスポート機能
4. エラーハンドリング

### テストデータ

`test/` ディレクトリにテスト用のCSVファイルを配置

## 問題報告

バグや改善提案は [GitHub Issues](https://github.com/make-it-at/ebaytool_sell/issues) で報告してください。

### 報告時の情報

- 発生環境（ブラウザ、OS）
- 再現手順
- 期待される動作
- 実際の動作
- エラーメッセージ（あれば）

## 質問・相談

開発に関する質問は GitHub Discussions または Issues でお気軽にお声がけください。 