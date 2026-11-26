# Aurora Serverless v2 (MySQL 8.4.7) A5:SQL Mk-2 CSV一括出力・性能試験データ分析ツール

Windows 11 の **PowerShell 5.1 標準機能のみ** で動作する、A5:SQL Mk-2 登録データベース（EC2踏み台 SSHトンネル経由 Aurora Serverless v2 / MySQL 8.4.7）からの自動CSV一括エクスポートおよび性能試験向けCSV総合データ分析・Excelレポート生成ツールです。

## ファイル構成

| ファイル | 役割 |
| :--- | :--- |
| `Run_AuroraCsvAnalyzer.bat` | ダブルクリック起動用ランチャー（PowerShell 5.1 `-STA` モードで `.ps1` を呼び出し） |
| `AuroraCsvAnalyzer.ps1` | 本体スクリプト（GUIダイアログ、A5M2設定読取、OpenSSHトンネル、純.NET MySQL 8.4.7クライアント、高速CSV分析エンジン、Excelレポート生成） |
| `tests\Verify-All.ps1` | 自動検証テストスイート（構文・MySQL 8.4.7プロトコル・CSV分析・Excel COM/OpenXML色/フォント/グラフ検証） |

## 動作要件

- **OS**: Windows 11
- **実行環境**: Windows PowerShell 5.1（標準搭載、追加モジュール・外部ツール不要）
- **SSHトンネル**: Windows 11 標準 OpenSSH Client (`C:\Windows\System32\OpenSSH\ssh.exe`)
- **Excel出力**: Microsoft Excel がインストールされている場合は Excel COM を使用してリッチなグラフ付き `.xlsx` を生成。未インストール環境でも .NET 標準機能による OpenXML (`.xlsx`) フォールバックで出力可能。

## 使い方（GUIモード：ダブルクリック起動）

1. `Run_AuroraCsvAnalyzer.bat` をダブルクリックします。
2. ダイアログが表示されます。目的に応じてモードを選択します：
   - **登録DBから直接分析Excel生成 (CSV出力なし)**: A5:SQL Mk-2 登録DBの接続設定（EC2踏み台SSHトンネル・SSL含む）経由で対象DBの全 `BASE TABLE` をストリーミング読込し、**ローカルへCSVファイルを一切出力せずに** 直接8シート構成の分析Excelレポートを生成します。CSV保存フォルダの指定は不要です。
   - **A5M2登録DBからCSV一括出力 ＋ 分析Excel生成**: 登録DB名を選ぶと、Aurora ホスト、ポート、スキーマ、ユーザー、SSHトンネル、SSL設定を読み込みます。既定の読込元はレジストリ (`HKCU:\Software\mmatsubara\A5M2*\DataSources`) です。
   - **既存CSVフォルダを分析してExcel生成**: すでにA5:SQL Mk-2等で一括エクスポートしたCSVフォルダを指定して分析レポートのみを生成します。
   - **CSV一括出力のみ**: 対象DBの全 `BASE TABLE` をCSVファイルとして一括エクスポートします。
   - **テスト前後の分析ファイル2件を比較してExcel生成**: テスト開始前に取得した分析Excel (`.xlsx`) とテスト開始後に取得した分析Excel (`.xlsx`) の2ファイルを指定し、更新のあったテーブル・追加のあったテーブル・変更のなかったテーブルおよびカラム/統計差分をまとめた比較Excelレポートを出力します（DB直接分析で出力したファイルもそのまま比較可能）。
3. **入力 / CSV保存フォルダ**（直接分析モード時は不要、比較モード時は **テスト開始前 / 開始後の分析ファイル2件**）と **出力先Excel (.xlsx)** を参照ボタンから指定し、**実行開始** をクリックします。

DBパスワードや SSH 秘密鍵のパスフレーズは実行時に入力します。DB接続時（CSV一括出力・DB直接分析）には接続セッションに `SET SESSION TRANSACTION READ ONLY` を設定します。

### 処理の中断・完全終了と、出力済みCSVからの続行（レジューム）

1. **中断する**には、GUIの **［中断］** またはCLIの **Ctrl+C** を押します。処理中の接続を閉じ、書きかけの `.csv.part` を削除します。CLIの中断時の終了コードは `130` です。GUIの **［閉じる］** では中断処理のあとに画面も閉じます。
2. **再開する**には、CSV出力を伴うモードで前回と同じ保存フォルダを指定します。GUIは中断記録を見つけると **［出力済みCSVから続行］** を選択するので、そのまま **［実行開始］** を押します。CLIでは同じ `-InputFolder` に `-Resume` を付けます。
3. **再開単位はテーブル**です。完了済みテーブルは保存済みCSVを使い、中断したテーブルは先頭から取得します。Excel作成中に中断した場合も、続行を選べば完成したCSVからレポートを作成できます。再開後に取得するテーブルには、その時点のDBの内容が反映されます。

CSV、進捗記録 `_export_progress.json`、分析Excelは、それぞれ書き込みを完了してから同じフォルダ内で置換します。再開時は記録されたファイル名とバイト数を照合し、未記録・サイズ不一致・未完了のテーブルを再取得します。同じフォルダへの同時出力はできません。

再開には前回と同じホスト、ポート、スキーマ、行数上限を指定してください。旧版の進捗記録も読み込めますが、旧版にはポートの記録がありません。進捗記録のない旧フォルダで明示的に続行した場合は、同名の空でないCSVを完了済みとみなすため、同じDBから出力を完了したファイルを使用してください。

### CSV出力中の進捗と高速化

GUIには処理全体とCSV出力の進捗率が表示されます。実行ログとCLIには対象テーブル名、CSV出力率、現在の出力行数が表示されます。各テーブルを同じ重みとし、処理中のテーブル内の割合はMySQLの推定行数から計算します。推定行数が不明な場合は完了テーブル数を使い、現在の出力行数を併記します。CSV出力の100%は全対象の確定後に表示します。[MySQL 8.4 の TABLE_ROWS の説明](https://dev.mysql.com/doc/refman/8.4/en/information-schema-tables-table.html)

CSV出力＋分析モードはDB通信をバッファリングし、新規テーブルの行をCSVに書きながら統計を集計します。再開時は前回完了したテーブルのCSVだけを読み直します。レポートの対象は、今回の出力または再開で選択されたテーブルです。

### ポータブルモードの登録DBを使う

1. A5:SQL Mk-2 で使用するポータブルワークスペースを開き、［データベース］→［データベースの追加と削除］→［エクスポート］で DBリスト (`.a5dblist` または `.dblist`) を保存します。ポータブル設定本体の `Setting.cini` は暗号化されています。[開発元の説明](https://a5m2.mmatsubara.com/wp/?topic=データベースツリーが表示されない)
2. 本ツールで DB接続を伴うモード（直接分析・CSV一括出力＋分析・CSV一括出力のみ）を選び、［ポータブルDBリスト選択...］で保存したファイルを開いて、［A5M2 登録DB名］から接続先を選びます。そのファイルに含まれる登録DBだけが表示されます。
3. DBパスワードと、必要に応じて SSH 秘密鍵のパスフレーズを入力し、［実行開始］をクリックします。接続先を変更したときは A5:SQL Mk-2 から再度エクスポートし、本ツールの［再読込］で反映します。［登録DBの読込元］でレジストリにも切り替えられます。

DBリスト内の SSH 秘密鍵と CA 証明書の相対パスは、DBリストを保存したフォルダを基準に解決します。

## コマンドライン（CLI / 自動化）利用例

```powershell
# A5:SQL Mk-2登録DBからCSVを出力せず、直接分析Excelレポートを生成
.\Run_AuroraCsvAnalyzer.bat -Mode DirectAnalyze -NoGui -A5DbName "aurora-prod-mysql" -DbPassword "your_pass" -OutputPath "C:\work\direct_report.xlsx"

# ポータブル版DBリストの登録DB名を指定して、CSV未出力で直接分析Excelレポートを生成
.\Run_AuroraCsvAnalyzer.bat -Mode DirectAnalyze -NoGui -A5DbListPath "E:\A5 SQL\接続先.a5dblist" -A5DbName "検証用Aurora" -DbPassword $env:AURORA_DB_PASSWORD -OutputPath "C:\work\direct_report.xlsx"

# 既存CSVフォルダの分析のみ
.\Run_AuroraCsvAnalyzer.bat -Mode AnalyzeOnly -NoGui -InputFolder "C:\work\exported_csv" -OutputPath "C:\work\report.xlsx"

# A5:SQL Mk-2登録DB名を指定して自動CSV一括エクスポート ＋ 分析実行
.\Run_AuroraCsvAnalyzer.bat -Mode ExportAndAnalyze -NoGui -A5DbName "aurora-prod-mysql" -DbPassword "your_pass" -InputFolder "C:\work\exported_csv" -OutputPath "C:\work\report.xlsx"

# 中断したフォルダの出力済みCSVをスキップして続きからエクスポート ＋ 分析Excel生成
.\Run_AuroraCsvAnalyzer.bat -Mode ExportAndAnalyze -NoGui -Resume -A5DbName "aurora-prod-mysql" -DbPassword "your_pass" -InputFolder "C:\work\exported_csv" -OutputPath "C:\work\report.xlsx"

# ポータブル版からエクスポートしたDBリストと登録DB名を指定してCSV一括出力（-Resume で途中再開も可能）
.\Run_AuroraCsvAnalyzer.bat -Mode ExportOnly -NoGui -Resume -A5DbListPath "E:\A5 SQL\接続先.a5dblist" -A5DbName "検証用Aurora" -DbPassword $env:AURORA_DB_PASSWORD -InputFolder "C:\work\exported_csv"

# テスト開始前・開始後の分析Excelファイル2件を比較して比較Excelレポートを生成
.\Run_AuroraCsvAnalyzer.bat -Mode CompareAnalysis -NoGui -BeforeAnalysisPath "C:\work\before_analysis.xlsx" -AfterAnalysisPath "C:\work\after_analysis.xlsx" -OutputPath "C:\work\diff_report.xlsx"
```

ポータブル版の例では、実行前に環境変数 `AURORA_DB_PASSWORD` に DBパスワードを設定します。`-A5DbListPath` の指定時はそのファイルだけから登録DBを検索します。`-Mode DirectAnalyze` 指定時は `-InputFolder` は不要で、`-OutputPath` を省略した場合はカレントディレクトリに `Aurora_PerfAnalysis_<日時>.xlsx` を出力します。`-Resume` を指定すると、`-InputFolder` 内の出力済み有効CSVをスキップし、残りのテーブルのみをエクスポートしてから後続処理（分析・Excel出力）を実行します。

## 出力Excelレポートの構成とデザイン仕様

- **フォント**: 全シート・全グラフで **Meiryo UI** を統一使用
- **配色（モノトーンベース）**:
  - 本文・軸・ラベル文字色: 濃いグレー (`#333333`)
  - アクセントカラー（タイトル・テーブルヘッダー・主要系列）: ミッドナイトネイビー (`#1A2B4C`)
  - 警告・注意文字色（高リスク・空テーブル・ホットキー偏り・孤立FK等）: バーガンディ (`#800020`)
- **8シート構成（グラフ付き）**:
  1. `01_サマリー`: KPIカード、テーブル別規模TOP15、レコード数横棒グラフ、容量構成比ドーナツグラフ、Aurora Serverless v2総合所見
  2. `02_テーブル一覧`: 全CSVの文字コード・行数・割合・行長・推定InnoDB容量・推定役割・警告一覧
  3. `03_カラム分析`: 型推定、NULL率、カーディナリティ・INDEX選択性、UTF-8バイト長(P95/最大)、Top1/Top3占有率、ジニ係数、偏りリスク＋ワーストTOP12グラフ
  4. `04_分布・割合分析`: カテゴリ・ステータス・区分値の件数/割合/累積割合（JMeterシナリオ重み付け用）＋上位カテゴリ構成比グラフ
  5. `05_数値統計・ヒストグラム`: Min / P25 / P50 / 平均 / P75 / P90 / P95 / P99 / Max / 標準偏差 / ゼロ率 / 負値率 / 外れ値率 ＋ 10区間ヒストグラム＆グラフ
  6. `06_時系列分析`: 日時期間・日次ピーク倍率・月別推移・時間帯別(00〜23時)分布・曜日別分布 ＋ 3種時系列グラフ
  7. `07_リレーション・ファンアウト`: 親子テーブルFK自動推定、参照一致率、孤立レコード数、1:Nファンアウト倍率（平均/P50/P95/P99/最大）＋比較グラフ
  8. `08_性能試験ガイド・警告`: 重要度別の性能リスク・警告一覧（バーガンディ強調）、JMeter CSV Data Set並列適性評価

### テスト前後比較Excelレポートの構成（6シート・グラフ付き）

単体分析レポートと同一のモノトーンデザイン仕様（**Meiryo UI** / 本文 `#333333` / アクセント `#1A2B4C` / 警告・注意 `#800020`）で、以下の6シートを出力します（Excel COM / OpenXML フォールバック両対応、COM出力ファイル・OpenXML出力ファイルの混在比較にも対応）。

1. `01_比較サマリー`: KPIカード（比較対象総数・更新あり・追加あり・変更なし・総レコード増減・要注意テーブル）、重複なし変更区分別サマリー表、レコード増減TOP15、グラフ（変更区分構成ドーナツ・増減数横棒）、総合所見
2. `02_全テーブル比較一覧`: 全テーブルのテスト前後レコード数・増減/増減率・列数・CSV/InnoDB容量増減・変化列数・主な変更内容・警告一覧
3. `03_更新のあったテーブル`: 既存データ更新（件数変動なしでカラム統計・カテゴリ分布・数値統計・日時統計・サイズ変化）・レコード減少（要注意・バーガンディ強調）・追加＋更新併用テーブル一覧（削除/消失テーブルがある場合は要注意セクションも併記）＋変更カラム数グラフ
4. `04_追加のあったテーブル`: 新規追加テーブルおよびレコード増加テーブル一覧（追加種別・開始前後件数・追加件数・増加率・容量増減・合計行）＋追加レコード数TOP15グラフ
5. `05_変更のなかったテーブル`: テスト開始前後で件数・サイズ・全カラム統計・分布・数値・日時統計が一致したテーブル一覧
6. `06_カラム・統計差分詳細`: 変更検出の根拠となったスキーマ変更・テーブル指標・カラム統計・カテゴリ分布・数値統計・時系列統計の開始前/開始後差分詳細一覧
