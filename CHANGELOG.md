# eBay出品作業効率化ツール 変更履歴

## [v1.5.14] - 2025-06-18
### パフォーマンス改善
- NGワードフィルタリング処理の高速化：処理時間を大幅に短縮
- 処理順序の最適化：CSVインポート完了後にフィルター処理を実行するよう修正
- エラーハンドリングの強化：データ不足時のエラーメッセージを改善

## [v1.5.4] - 2025-06-16
### 機能改善
- 全処理一括実行の処理順序とエラーハンドリングを改善
- データ存在チェックの追加：データがない場合は処理を中止
- カラム存在チェックの追加：必要なカラムがない場合は該当処理をスキップ
- 警告メッセージの強化：処理がスキップされた場合に詳細を表示

## [v1.5.3] - 2025-06-15
### パフォーマンス改善
- データ処理の高速化: バッチ処理の実装とメモリ使用量の最適化
- CSVインポート処理の最適化: 大量データの処理速度を向上
- NGワードフィルター処理の高速化: バッチ処理によるパフォーマンス改善
- キャッシュ機能の実装: 頻繁に使用するデータ構造をキャッシュ化

## [v1.5.2] - 2025-05-31
### UI改善
- 処理結果表示方法の変更: トースト通知を使用せず、サイドバーの結果メッセージのみに表示するよう変更
- ユーザーの視線を一箇所に集中させる設計に変更し、操作性を向上
- サイドバーのメッセージ表示スタイルを改善

## [v1.5.1] - 2025-05-30
### UI改善
- 処理結果表示の強化: タイムスタンプを追加
- サイドバーの結果メッセージにログ形式の日時表示を追加
- 全ての処理結果表示を統一形式に変更
- 結果メッセージのスタイルを改善し視認性を向上

## [v1.5.0] - 2025-05-29
### UI改善
- フィルター処理の結果表示を改善: 処理前後のデータ数比較を追加
- 全処理のメッセージで「○件 → ○件」の形式でデータ数変化を表示
- NGワード、重複、文字数、価格フィルターなど全ての処理結果画面を統一
- サイドバーでの結果表示を強化し、より詳細な処理情報を視覚的に表示

## [v1.4.9] - 2025-05-27
### 仕様修正
- CSVインポート処理の仕様を変更: 「データインポート」シートへのインポートを廃止し、「出品データ」シートに直接インポートする方式に統一
- 「データインポート」シートへの参照を全て削除し、全フィルター処理が「出品データ」シートのみを対象とするように修正
- シート構造をシンプル化し、処理フローを改善

## [v1.4.8] - 2025-05-26
### 機能改善とUI改善
- NGワードフィルター機能の強化: 大文字・小文字・スペースの違いを無視する機能を追加
- 検索用に文字列を正規化する関数を追加（小文字変換、連続スペースの単一化など）
- ヘルプドキュメントの拡充: NGワード設定に関する新機能の説明を追加
- 処理結果メッセージの表示改善: 削除件数・修正件数などの詳細を表示

## [v1.4.6] - 2025-05-23
### UI改善とバグ修正
- ボタンのサイズとスタイルを調整し、テキストが収まるように修正
- ボタンの高さを44pxに統一し、適切なパディングとwhite-space設定を追加
- CSVインポート時の「Cannot read properties of null (reading 'clearContents')」エラーを修正
- シートが存在しない場合に自動作成する機能を追加

## [v1.4.5] - 2025-05-22
### UI改善
- サイドバーUIの改善：ヘッダータイトルを削除し、シンプルなバージョン表示に変更
- すべてのボタンサイズを統一（高さ40px、幅100%）に標準化し、一貫性を向上
- 個別フィルター処理ボタンの間隔調整

## [v1.4.4] - 2025-05-21
### UI改善
- サイドバーUIのセクション名称変更：「データ処理」→「CSV入出力」
- 「NGワード管理」セクションを「設定」に名称変更し、最下部に配置
- 不要だった「設定」セクションを削除
- アプリケーション名を「みずのとい」に変更
- フッター著作権表示を更新

## [v1.4.3] - 2025-05-20
### UI改善
- サイドバーUI改善：全処理一括実行を最上部に配置し、個別フィルターをトグルで整理
- インポートダイアログボタンを削除し、サイドバーインターフェースを簡素化
- 使用頻度の高い操作を優先した配置に変更

## [v1.4.2] - 2025-05-19
### UI改善
- メニューを簡素化し、「サイドバーを表示」と「ヘルプ」のみに整理
- 機能をすべてサイドバーからアクセスするように変更

## [v1.4.1] - 2025-05-18
### 改善
- CSVエクスポート機能の簡素化：ワンクリックで出品データをエクスポート
- シート選択プロセスを削除し、直接出品データシートをエクスポート

## [v1.4.0] - 2025-05-17
### 機能追加
- CSVエクスポート機能の改善：出品データシートとデータインポートシートの両方からエクスポート可能に
- サイドバーからの直接CSVダウンロード機能を追加
- エクスポート時のシート選択機能を追加

### 不具合修正
- CSVエクスポート中の「Cannot read properties of null (reading 'getDataRange')」エラーを修正
- シートが存在しない場合のエラーハンドリングを改善

## [v1.3.9] - 2025-05-16
### 不具合修正
- 全処理一括実行時の処理時間が「不明」と表示される問題を修正
- 各フィルター処理完了後も全処理一括実行の処理時間を正しく計測できるよう改善

## [v1.3.8] - 2025-05-15
### 不具合修正
- NGワードフィルタリング完了メッセージでの「Cannot read properties of undefined (reading '4')」エラーを修正
- 部分削除処理の件数計算ロジックを改善し、エラーハンドリングを追加

## [v1.3.7] - 2025-05-15
### 不具合修正
- NGワードフィルタリング実行時の「Cannot read properties of undefined (reading '4')」エラーを修正
- 設定シートのヘッダー行の存在確認を追加
- 設定値取得処理を改善し、各設定値取得時のエラーハンドリングを強化

## [v1.3.6] - 2025-05-14
### 変更
- 設定シートの形式を変更：NGワード処理を「リスト削除」と「削除ワード」の2列に分離
- NGワードフィルタリング機能を修正：モード選択不要の2段階処理に変更
  - 「リスト削除」列のNGワードを含むリストを削除
  - 「削除ワード」列のNGワードの部分のみを削除
- 文字数制限と価格下限を設定シートから直接取得する方式に変更

## [v1.3.5] - 2025-05-14
### 修正
- 設定値が取得できない場合はデフォルト値で処理せず、エラーメッセージを表示して処理を中止するように変更
- 文字数フィルターと価格フィルターの両方に設定値チェックを追加

## [v1.3.4] - 2025-05-14
### 修正
- 文字数フィルターと価格フィルターの条件を「以下」に修正
- 設定値が読み込めない問題を修正（デフォルト値を適用）
- フィルター処理のログメッセージを改善

## [v1.3.3] - 2025-05-14
### 修正
- 文字数フィルターの条件を「未満」に修正（以下→未満）
- 価格フィルターの条件を「未満」に修正（以下→未満）
- 価格フィルターで文字列を適切に数値変換するよう改善
- フィルター結果メッセージに具体的な条件値を表示

## [v1.3.2] - 2025-05-14
### 修正
- 重複チェック機能を完全一致のみに変更
- 類似度による判定からタイトルの完全一致による判定に修正

## [v1.3.1] - 2025-05-13
### 修正
- フィルター処理の各ボタンがクリックに反応しない問題を修正
- Filters名前空間の関数をグローバル関数としてエクスポートし、サイドバーから直接呼び出せるように変更

## [v1.3.0] - 2025-05-13
### 変更
- 全体的なコード構造を改善
- フィルター処理機能の効率化
- UI/UXの向上

## [v1.0.6] - 2025-05-03
### 変更
- CSVインポート処理を改善し、データを加工せずそのまま忠実にインポートするよう変更
- ヘッダーマッピング処理を削除し、CSVをそのまま表示するよう修正

## [v1.0.5] - 2025-05-02
### 修正
- 日本語を含むCSVファイルのエンコード問題を修正
- Unicode文字のBase64エンコードのためにencodeURIComponent()とunescape()を使用

## [v1.0.4] - 2025-05-01
### 修正
- CSVインポート処理の問題を修正
- Base64エンコードをクライアント側で行い、importCsvFromBase64関数を追加

## [v1.0.3] - 2025-04-30
### 修正
- UI名前空間が認識されない問題を修正
- サイドバーからUI関数を直接呼び出せるようにApp.gsにグローバル関数を追加

## [v1.0.2] - 2025-04-29
### 修正
- スタイルが適用されない問題を修正
- インラインスタイルを直接HTMLに記述し、Font AwesomeをCDNから読み込む方式に変更

## [v1.0.1] - 2025-04-28
### 修正
- 「importCsv is not defined」エラーを修正
- onclick属性からeventListener方式に変更し、window.importCsv形式で関数を定義

## [v1.0.0] - 2025-04-27
### 追加
- 初回リリース
- CSVインポート/エクスポート機能
- NGワードフィルタリング機能
- 重複チェック機能
- 文字数フィルター機能
- 所在地情報修正機能
- 価格フィルター機能

## v1.2.0 - 2024-07-16

### 改善
- バージョン管理システムの改善
- NGワードフィルタリング機能の問題修正完了を確認
- Title列の参照を正確に行えるように修正

## v1.1.9 - 2024-07-15

### 追加
- 「出品データ」シートの追加
  - テスト用CSVインポート時に「出品データ」シートに直接転記する機能
  - 「データインポート」シートとは別に出品用データを管理
- フィルター機能のグローバルエントリーポイントを追加
  - 全てのフィルター機能（NGワード、重複チェック、文字数制限、所在地情報修正、価格フィルター）に直接アクセスできるグローバル関数を追加
  - サイドバーからの実行時のエラーを修正

### 修正
- テスト用CSVインポート機能のターゲットを「データインポート」から「出品データ」シートに変更
- テストデータヘルプダイアログの説明文を更新
- 各フィルター機能（NGワード、重複チェック、文字数制限、所在地情報修正、価格フィルター）を修正し、「出品データ」シートを使用するように変更

## v1.1.8 - 2024-07-15
### 追加
- 同一フォルダ内のCSVファイル直接読み込み機能
  - テスト用CSVインポート時に「eBay Export May 12 2025.csv」を自動的に検索して読み込み
  - プロジェクトと同じフォルダおよびDrive内を検索
- テストデータ管理機能
  - 専用の「テストデータ」シートを追加
  - テストデータ設定方法を説明するヘルプダイアログを追加
  - ローカルパス（/Users/kyosukemakita/Documents/Cursor/ebaytool_sell/eBay Export May 12 2025.csv）のテストデータを取り込むためのサポート

### 変更
- テスト用CSVインポート機能の改善
  - ハードコードされたデータからテストデータシートを読み込む方式に変更
  - エラーハンドリングの強化
  - テストデータシートから「データインポート」シートへのデータ転記プロセスの最適化

## v1.1.7 - 2024-07-12
### 修正
- テスト用CSVインポート機能の改善
- Google Driveに依存せず、ハードコードされたテストデータを使用するように変更
- 5件のサンプルデータを内蔵

## v1.1.6 - 2024-07-12
### 機能追加
- テスト用CSVインポート機能を追加
- 「eBay Export May 12 2025.csv」ファイルを直接インポートするボタンをサイドバーに追加

## v1.1.5 - 2024-07-12
### 修正
- サイドバーのバージョン表示を修正
- バージョン情報がテンプレートに正しく渡されるように改善

## v1.1.4 - 2024-07-12
### 改善
- 通知方法を改善し、ダイアログではなくスプレッドシートのトーストとサイドバーで表示するように変更
- 不要なモードレスダイアログの完全な排除

## v1.1.3 - 2024-07-12
### 改善
- CSVエクスポート機能を改善し、Google Drive保存ではなくローカルPCに直接ダウンロードするように変更
- 不要なダイアログ表示を削除し、サイドバーでの通知に統一

## v1.1.2 - 2024-07-11
### 改善
- 所在地情報修正機能を修正し、「所在地」または「Location」カラムを自動検出するように改善
- エディタから所在地情報修正機能を直接実行できるようにグローバル関数を追加
- サイドバーでの所在地情報修正ボタンのイベントハンドラを修正

## v1.1.1 - 2024-07-10
### 修正
- UIが利用できない環境でのエラーを修正
- バージョン表示の不整合を修正

## v1.1.0 - 2024-07-09
### 機能追加
- NGワードフィルタリング機能を実装
- 重複チェック機能を実装
- 文字数制限フィルタリング機能を実装
- 所在地情報修正機能を実装
- 価格フィルタリング機能を実装
- バージョン管理機能を追加

## v1.0.0 - 2024-07-08
### 初期リリース
- CSVインポート/エクスポート機能
- 基本的なUI実装
- 設定管理機能 