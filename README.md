# Apache JMeter エンドツーエンド負荷テスト実践パッケージ

Windows 11 開発端末でのシナリオ作成から、Linux 実行サーバでの Non-GUI (CLI) 負荷テスト実行、リアルタイム監視、JTL ログ回収、および Excel (Meiryo UI / モノトーン基調) による性能分析レポート作成までを完全網羅した実務運用パッケージです。

---

## 📁 ディレクトリ構成

```text
JMeter_Tutorial_Agy/
├── JMeter_EndToEnd_Guide_And_Report.xlsx   # 【本体】マニュアル & 性能分析レポート (Excel)
├── README.md                               # 本ドキュメント
├── build_jmeter_workbook.py                # Excelブック再生成用 Python スクリプト
├── scenarios/
│   ├── sample_test_plan.jmx                # 標準ECサイトWeb API負荷テストシナリオ
│   └── sample_cognito_apigw.jmx            # 【NEW】Cognito OAuth2認証 & API Gateway連携シナリオ
├── data/
│   └── sample_users.csv                    # パラメータ検証用CSVデータ
├── scripts/
│   ├── workbook/                           # Excelシート生成モジュール群 (全9シート)
│   ├── setup_windows_jmeter.ps1            # Windows 11 環境構築支援スクリプト
│   ├── setup_linux_jmeter.sh               # Linux 実行サーバ環境構築スクリプト
│   ├── run_load_test.sh                    # Linux CLI 負荷テスト自動実行スクリプト
│   └── stop_test.sh                        # 安全停止 (Graceful Shutdown) スクリプト
├── results/
│   └── sample_run.jtl                      # 検証用サンプル JTL (CSV形式)
└── logs/
    └── sample_jmeter.log                   # 検証用サンプル JMeter ログ
```

---

## 📊 Excel ワークブックの構成 (`JMeter_EndToEnd_Guide_And_Report.xlsx`)

全シートにおいて **フォント: Meiryo UI**、**配色: モノトーン基調 (スレートグレー / チャコール)**、**破線枠による画像貼付エリア (スクリーンショット枠)** を完備しています。

| No | シート名 | 主な内容 | 画像貼付枠 |
|:---|:---|:---|:---:|
| 00 | `00_表紙・全体概要` | 文書管理情報、End-to-End ワークフロー全体図、各シート目次 (全9シート) | システム構成図・配置図 |
| 01 | `01_Win環境構築` | Java 17/21、JMeter 5.6.3、環境変数設定、setenv.bat、Plugins Manager導入、主要プラグイン一覧 | 環境変数設定画面、Plugins Manager画面 |
| 02 | `02_シナリオ作成手法` | テスト計画、スレッドグループ、HTTP初期値、CSV設定、アサーション、画面実行手順、リスナー確認観点、GUI実行禁止則 | スレッドグループ設定、デバッグ成功画面 |
| 03 | `03_Cognito・APIGW連携` | 【NEW】Cognitoトークン取得、BASE64エンコード関数、正規表現抽出でのToken保持、X-API-IDヘッダ、API連携 | Cognito/正規表現設定画面、APIGW疎通証跡 |
| 04 | `04_Linux実行基盤構築` | 推奨ディレクトリ設計、Java/JMeter配備、JVMヒープ設定、OSカーネルチューニング、SCP/WinSCP転送 | 転送実行画面、Linux配置確認画面 |
| 05 | `05_CLI実行と監視` | Non-GUIコマンド構文 (-n -t -l -j -e -o)、Summariser読解、バックグラウンド実行、安全停止 | CLI実行・進捗画面、top/vmstat監視画面 |
| 06 | `06_結果取得とExcel取込` | JTL全17フィールド定義、Linuxから回収、Power Query取込手順、エポックミリ秒→JST変換数式 | SCP結果回収画面、Power Query取込画面 |
| 07 | `07_性能分析レポート` | 【報告書】総合KPIカード、SLA達成判定表、トランザクション別統計表(Excel数式)、エラー分析、提言 | 応答時間推移、TPS推移、インフラ監視 |
| 08 | `08_確認観点と注意点` | プロが確認する14の必須観点・落とし穴チェックリスト (Think Time、Coordinated Omission等) | 事前確認印・レビュー承認証跡 |

---

## 🚀 クイックスタートガイド

### 1. Windows 11 側での準備
```powershell
# PowerShell を管理者として実行
cd C:\Users\taka_\Agy\JMeter_Tutorial_Agy
# 環境確認と setenv.bat 生成
.\scripts\setup_windows_jmeter.ps1
```

### 2. Linux 実行サーバ側での準備
```bash
# Linux サーバ上で実行 (Java 17 & JMeter 5.6.3 セットアップ)
bash scripts/setup_linux_jmeter.sh
```

### 3. シナリオ (JMX) と CSV の転送
```powershell
# Windows から Linux サーバへ転送
scp -P 22 .\scenarios\sample_cognito_apigw.jmx user@linux-host:/data/jmeter/scenarios/
scp -P 22 .\data\sample_users.csv user@linux-host:/data/jmeter/csv/
```

### 4. Linux サーバ上での CLI 実行
```bash
# 50ユーザー、60秒ランプアップ、1200秒 (20分間) 実行
/data/jmeter/scripts/run_load_test.sh 50 60 1200
```

### 5. 結果 (JTL) の回収と Excel 分析
```powershell
# Linux から Windows へ結果回収
scp -P 22 user@linux-host:/data/jmeter/results/*.jtl C:\PerformanceTest\Results\
```
回収した JTL ファイルを Excel の `07_性能分析レポート` に反映し、必要に応じてスクリーンショットを枠内に貼り付けて提出します。
