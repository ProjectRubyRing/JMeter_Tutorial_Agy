# -*- coding: utf-8 -*-
"""
Sheet 05: CLI実行と監視 (Non-GUI Execution & Monitoring)
"""

from openpyxl.utils import get_column_letter

def build_sheet_cli_execution(wb, styles):
    ws = wb.create_sheet(title="05_CLI実行と監視")
    ws.views.sheetView[0].showGridLines = True

    styles["apply_title_banner"](
        ws,
        "05. Linux CLI (Non-GUI) 負荷テスト実行・リアルタイム監視ガイド",
        "CLI構文完全解説・バックグラウンド実行・Summariser進捗監視・リソース監視・安全停止",
        10
    )

    cur_r = 4
    styles["apply_section_header"](ws, cur_r, "1. JMeter Non-GUI (CLI) 実行コマンドの構文と全オプション解説", 10)
    cur_r += 1

    cmd_box = (
        "【基本実行コマンド構文】\n"
        "jmeter -n -t /data/jmeter/scenarios/scenario.jmx \\\n"
        "       -l /data/jmeter/results/run_$(date +%Y%m%d_%H%M%S).jtl \\\n"
        "       -j /data/jmeter/logs/jmeter_$(date +%Y%m%d_%H%M%S).log \\\n"
        "       -e -o /data/jmeter/results/report_$(date +%Y%m%d_%H%M%S)/ \\\n"
        "       -Jthreads=50 -Jrampup=60 -Jduration=1200"
    )
    ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r + 2, end_column=10)
    for r in range(cur_r, cur_r + 3):
        for c in range(1, 11):
            cell = ws.cell(row=r, column=c)
            cell.fill = styles["fill_code"]
            cell.border = styles["cell_border"]
            if r == cur_r and c == 1:
                cell.value = cmd_box
                cell.font = styles["font_code"]
                cell.alignment = styles["align_left"]
    ws.row_dimensions[cur_r].height = 20
    ws.row_dimensions[cur_r + 1].height = 20
    ws.row_dimensions[cur_r + 2].height = 24
    cur_r += 4

    cli_opts = [
        ("-n", "Non-GUI モード指定", "GUI (画面) を起動せず、コンソールのみで軽量実行する必須フラグ。画面描画オーバーヘッドをゼロにします。", "本番負荷テストでは必ず指定。未指定だとGUI起動を試みてエラーになる。"),
        ("-t <jmx_path>", "テスト計画ファイル指定", "実行するシナリオファイル (.jmx) のパスを指定します。相対パスまたは絶対パスが有効。", "/data/jmeter/scenarios/sample_test_plan.jmx などを指定。"),
        ("-l <jtl_path>", "実行結果ログ出力先指定", "テスト結果 (JTLサンプリングデータ) を出力するCSV形式ファイルのパス。分析レポートの元データ。", "/data/jmeter/results/run_20261004.jtl などを指定。既存ファイルがある場合は追記されるため日付毎に新規命名を推奨。"),
        ("-j <log_path>", "JMeterシステムログ指定", "JMeter内部の動作ログ、エラーログ、例外スタックトレースを出力するファイルのパス。", "/data/jmeter/logs/jmeter.log などを指定。サンプラーのHTTPエラーではなくJMeter自体の異常を検知する。"),
        ("-e", "HTMLレポート生成トリガー", "テスト完了時に、自動的にグラフィカルなHTMLダッシュボードレポートを生成するフラグ。", "-o オプションとセットで必ず使用する。"),
        ("-o <dir_path>", "HTMLレポート出力先フォルダ", "生成されたHTMLダッシュボードを出力するディレクトリパス。", "※重要: 出力先ディレクトリは【空】であるか、存在しないパスである必要があります。既存ファイルがあるとエラー停止します。"),
        ("-J<prop>=<val>", "JMeterローカルプロパティ上書き", "シナリオ内のプロパティ (${__P(threads, 10)}) をCLI実行時に動的注入する。", "例: -Jthreads=100 -Jrampup=30。JMXファイルを書き換えることなく負荷強度を柔軟に変更可能。"),
        ("-D<prop>=<val>", "Javaシステムプロパティ定義", "JVM全体のシステムプロパティを設定する (例: SSL証明書の検証無効化など)。", "-Djsse.enableSNIExtension=true などのインフラ固有設定に利用。")
    ]

    tbl_headers5 = ["オプション", "名称・役割", "詳細機能・設定内容", "実務上の注意点・ベストプラクティス"]
    ws.row_dimensions[cur_r].height = 24
    cols_map5 = [(1, 2), (3, 4), (5, 7), (8, 10)]
    for i, (sc, ec) in enumerate(cols_map5):
        if sc != ec:
            ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_tbl_header"]
            cell.border = styles["header_border"]
            if c == sc:
                cell.value = tbl_headers5[i]
                cell.font = styles["font_tbl_header"]
                cell.alignment = styles["align_center"]

    cur_r += 1
    for idx, opt in enumerate(cli_opts):
        ws.row_dimensions[cur_r].height = 32
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # Option
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
        for c in range(1, 3):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = opt[0]
                c_cell.font = styles["font_code"]
                c_cell.alignment = styles["align_center"]

        # Name
        ws.merge_cells(start_row=cur_r, start_column=3, end_row=cur_r, end_column=4)
        for c in range(3, 5):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 3:
                c_cell.value = opt[1]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Description
        ws.merge_cells(start_row=cur_r, start_column=5, end_row=cur_r, end_column=7)
        for c in range(5, 8):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 5:
                c_cell.value = opt[2]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        # Practical Points
        ws.merge_cells(start_row=cur_r, start_column=8, end_row=cur_r, end_column=10)
        for c in range(8, 11):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 8:
                c_cell.value = opt[3]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]
        cur_r += 1

    # Section 2: Real-time Monitoring & Summariser Reading
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "2. コンソール実行状況の監視手法と進捗ログ (Summariser) の読解", 10)
    cur_r += 1

    summariser_sample = (
        "【Summariser 出力ログの例】\n"
        "summary +   500 in 00:00:10 =   50.0/s Avg:   180 Min:    45 Max:  1250 Err:     0 (0.00%) Active: 50 Started: 50 Finished: 0\n"
        "summary =  1500 in 00:00:30 =   50.0/s Avg:   192 Min:    40 Max:  2100 Err:     2 (0.13%)"
    )
    ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r + 1, end_column=10)
    for r in range(cur_r, cur_r + 2):
        for c in range(1, 11):
            cell = ws.cell(row=r, column=c)
            cell.fill = styles["fill_code"]
            cell.border = styles["cell_border"]
            if r == cur_r and c == 1:
                cell.value = summariser_sample
                cell.font = styles["font_code"]
                cell.alignment = styles["align_left"]
    ws.row_dimensions[cur_r].height = 20
    ws.row_dimensions[cur_r + 1].height = 24
    cur_r += 3

    sum_fields = [
        ("summary +", "直近インターバルの増分", "直前の30秒間（インターバル期間）のみの統計値。現在の瞬間的な負荷状況を表す。"),
        ("summary =", "テスト開始からの累積値", "テスト開始から現在までの全リクエストの通算集計値。総合的な平均性能を表す。"),
        ("500 in 00:00:10", "処理リクエスト数と経過時間", "過去10秒間に500件のリクエストを処理完了したことを示す。"),
        ("50.0/s", "スループット (Throughput / TPS)", "1秒あたりに処理されたリクエスト数。想定通りのスループットが出ているか確認。"),
        ("Avg: 180", "平均応答時間 (Average Response Time)", "リクエスト送信からレスポンス受信完了までの平均時間 (ミリ秒)。"),
        ("Min: 45 / Max: 1250", "最小・最大応答時間", "対象インターバル中の最短レスポンスと最長レスポンス (ミリ秒)。Maxの突出に注意。"),
        ("Err: 0 (0.00%)", "エラー件数とエラー率", "HTTPエラーまたはアサーション失敗の件数と全体に対する割合。0%であるべき。"),
        ("Active: 50", "アクティブスレッド数 (同時仮想ユーザー数)", "現在並行してリクエスト処理・待機を行っている仮想ユーザー数。目標値に達しているか確認。")
    ]

    for sf in sum_fields:
        ws.row_dimensions[cur_r].height = 22
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
        for c in range(1, 3):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_card"]
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = sf[0]
                c_cell.font = styles["font_code"]
                c_cell.alignment = styles["align_left"]

        ws.merge_cells(start_row=cur_r, start_column=3, end_row=cur_r, end_column=4)
        for c in range(3, 5):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_card"]
            c_cell.border = styles["cell_border"]
            if c == 3:
                c_cell.value = sf[1]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        ws.merge_cells(start_row=cur_r, start_column=5, end_row=cur_r, end_column=10)
        for c in range(5, 11):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_white"]
            c_cell.border = styles["cell_border"]
            if c == 5:
                c_cell.value = sf[2]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]
        cur_r += 1

    # Section 3: Safe Shutdown Procedure
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "3. 安全停止（Graceful Shutdown）と強制終了手順", 10)
    cur_r += 1

    stop_methods = [
        ("安全停止 (Graceful Shutdown) 【推奨】", 
         "# JMeterインストールディレクトリ配下のスクリプトを実行:\n"
         "/opt/jmeter/apache-jmeter-5.6.3/bin/shutdown.sh\n"
         "または実行サーバの4445ポートにシャットダウンシグナル送信", 
         "現在実行中のリクエストの完了を待ち、JTLログをフラッシュして安全にプロセスを終了します。データ破損が起きません。"),
        ("即時停止 (Immediate Stop)", 
         "/opt/jmeter/apache-jmeter-5.6.3/bin/stoptest.sh", 
         "実行中のリクエストを即座にソケット切断して終了します。通信エラーが記録される可能性があります。"),
        ("強制キル (Kill Process) 【非常時のみ】", 
         "kill -9 $(pgrep -f ApacheJMeter.jar)", 
         "プロセスがハングアップし、shutdown.sh を受け付けない場合のみ使用。JTLファイルの末尾が破損する恐れがあります。")
    ]

    for sm in stop_methods:
        ws.row_dimensions[cur_r].height = 36
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=3)
        for c in range(1, 4):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_card"]
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = sm[0]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        ws.merge_cells(start_row=cur_r, start_column=4, end_row=cur_r, end_column=6)
        for c in range(4, 7):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_code"]
            c_cell.border = styles["cell_border"]
            if c == 4:
                c_cell.value = sm[1]
                c_cell.font = styles["font_code"]
                c_cell.alignment = styles["align_left"]

        ws.merge_cells(start_row=cur_r, start_column=7, end_row=cur_r, end_column=10)
        for c in range(7, 11):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_white"]
            c_cell.border = styles["cell_border"]
            if c == 7:
                c_cell.value = sm[2]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]
        cur_r += 1

    # Image Area for Sheet 5
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "4. CLI実行・コンソール監視証跡貼付エリア", 10)
    cur_r += 1
    styles["apply_image_placeholder"](
        ws, cur_r, 1, cur_r + 8, 5,
        "JMeter Non-GUI (CLI) 実行進捗ログ (Summariser) 画面",
        "Linuxコンソール上で実行中の Summariser 出力ログ (Active threads, TPS, Avg, Err%) のリアルタイム表示キャプチャ"
    )
    styles["apply_image_placeholder"](
        ws, cur_r, 6, cur_r + 8, 10,
        "Linux サーバ OS リソース監視画面 (top / vmstat / dstat)",
        "テスト実行中における Linux 負荷生成サーバおよびターゲットサーバの CPU使用率、メモリ、ロードアベレージ監視キャプチャ"
    )
    cur_r += 9

    # Set column widths
    col_widths5 = [12, 14, 14, 14, 16, 16, 16, 16, 16, 12]
    for idx, w in enumerate(col_widths5, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
