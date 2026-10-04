# -*- coding: utf-8 -*-
"""
Sheet 01: Win環境構築 (Windows 11 Setup & Plugins Guide)
"""

from openpyxl.utils import get_column_letter

def build_sheet_win_setup(wb, styles):
    ws = wb.create_sheet(title="01_Win環境構築")
    ws.views.sheetView[0].showGridLines = True

    styles["apply_title_banner"](
        ws,
        "01. Windows 11端末 JMeterセットアップ＆機能追加ガイド",
        "Java環境整備・JMeter本体導入・ヒープ調整・Plugins Manager・各種プラグイン追加・起動確認",
        10
    )

    cur_r = 4
    styles["apply_section_header"](ws, cur_r, "1. Windows 11 環境構築手順（ステップ・バイ・ステップ）", 10)
    cur_r += 1

    win_steps = [
        ("1", "Java (JDK) の入手とインストール", 
         "JMeterはJavaアプリケーションのため、JDK 17または21 LTSが必要です。\n"
         "・Amazon Corretto 17 または Eclipse Temurin 17 (x64 msi) を公式サイトより入手。\n"
         "・インストーラを実行し、デフォルトパス (例: C:\\Program Files\\Amazon Corretto\\jdk17.0.x) に導入。",
         "・PowerShellで確認:\n  java -version\n・バージョン 17 または 21 が表示されれば成功。", "必須"),
        ("2", "環境変数 JAVA_HOME / PATH の設定",
         "Windowsの検索バーで「環境変数」を検索 →「システム環境変数の編集」を開く。\n"
         "・システム環境変数に新規作成: 変数名「JAVA_HOME」、変数値「C:\\Program Files\\Amazon Corretto\\jdk17.0.x」\n"
         "・Path変数を編集し、末尾に「%JAVA_HOME%\\bin」を追加して最上位付近に移動。",
         "・PowerShellを新規に開き実行:\n  echo $env:JAVA_HOME\n設定したパスが返ればOK。", "必須"),
        ("3", "JMeter 本体のダウンロードと配置",
         "Apache JMeter 公式サイト (https://jmeter.apache.org/download_jmeter.cgi) より、\n"
         "「Binaries」項目の「apache-jmeter-5.6.3.zip」をダウンロード。\n"
         "・zipを展開し、日本語やスペースを含まない短めのパス (推奨: C:\\tools\\apache-jmeter-5.6.3) に配置。\n"
         "※注意: Program Files配下に置くと権限エラーや空白パスによるスクリプト障害の原因となるため避ける。",
         "・C:\\tools\\apache-jmeter-5.6.3\\bin\\jmeter.bat が存在することを確認。", "必須"),
        ("4", "環境変数 JMETER_HOME の設定",
         "システム環境変数に新規追加:\n"
         "・変数名: JMETER_HOME\n"
         "・変数値: C:\\tools\\apache-jmeter-5.6.3\n"
         "・Path変数の編集: 末尾に「%JMETER_HOME%\\bin」を追加。",
         "・任意のフォルダからPowerShellで:\n  jmeter -v\nを実行し、JMeterのバージョンロゴが表示されること。", "推奨"),
        ("5", "ヒープメモリ設定の最適化 (setenv.bat)",
         "シナリオ作成や単体検証時のOutOfMemoryErrorを防ぐためヒープを確保します。\n"
         "・C:\\tools\\apache-jmeter-5.6.3\\bin フォルダ直下に「setenv.bat」を新規作成。\n"
         "・以下を記述して保存 (UTF-8またはANSI):\n"
         "  set HEAP=-Xms1g -Xmx2g -XX:MaxMetaspaceSize=512m\n"
         "※setenv.batを作成すると、jmeter.bat起動時に自動で読み込まれます。",
         "・jmeter.bat 実行時のコンソール画面でヒープ引数が反映されていることを確認。", "必須"),
        ("6", "JMeter Plugins Manager の導入",
         "追加機能（各種スレッドグループ、グラフ、関数）をGUI上からワンクリック導入する必須管理機構です。\n"
         "・https://jmeter-plugins.org/wiki/PluginsManager/ より「jmeter-plugins-manager-1.10.jar」を入手。\n"
         "・ダウンロードしたjarを「C:\\tools\\apache-jmeter-5.6.3\\lib\\ext」フォルダ内に直接配置。\n"
         "・JMeterを再起動すると、上部メニュー「Options (オプション)」に「Plugins Manager」が追加される。",
         "・Options > Plugins Manager を開き、プラグイン一覧が表示されれば導入成功。", "強く推奨"),
        ("7", "JMeter GUI の起動確認と日本語/英語設定",
         "bin\\jmeter.bat をダブルクリックしてGUIを起動します。\n"
         "・GUI言語を英語に固定する場合 (現場推奨: ネット情報や公式ドキュメント、エラーログと完全一致するため):\n"
         "  bin\\jmeter.properties を開き、以下を修正:\n  language=en\n"
         "・日本語表示にする場合:\n  Options > Choose Language > Japanese を選択。",
         "・GUI画面が正常に立ち上がり、エラーダイアログが出ないこと。", "必須")
    ]

    tbl_headers2 = ["No", "構築工程", "詳細手順・設定内容", "検証コマンド・確認基準", "重要度"]
    ws.row_dimensions[cur_r].height = 24
    cols_map2 = [(1, 1), (2, 3), (4, 7), (8, 9), (10, 10)]
    for i, (sc, ec) in enumerate(cols_map2):
        if sc != ec:
            ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_tbl_header"]
            cell.border = styles["header_border"]
            if c == sc:
                cell.value = tbl_headers2[i]
                cell.font = styles["font_tbl_header"]
                cell.alignment = styles["align_center"]

    cur_r += 1
    for idx, step in enumerate(win_steps):
        ws.row_dimensions[cur_r].height = 54
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # No
        ws.cell(row=cur_r, column=1, value=step[0]).font = styles["font_body_bold"]
        ws.cell(row=cur_r, column=1).fill = fill_curr
        ws.cell(row=cur_r, column=1).border = styles["cell_border"]
        ws.cell(row=cur_r, column=1).alignment = styles["align_center"]

        # Name
        ws.merge_cells(start_row=cur_r, start_column=2, end_row=cur_r, end_column=3)
        for c in range(2, 4):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 2:
                c_cell.value = step[1]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Detail
        ws.merge_cells(start_row=cur_r, start_column=4, end_row=cur_r, end_column=7)
        for c in range(4, 8):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 4:
                c_cell.value = step[2]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        # Validation
        ws.merge_cells(start_row=cur_r, start_column=8, end_row=cur_r, end_column=9)
        for c in range(8, 10):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 8:
                c_cell.value = step[3]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        # Priority
        p_cell = ws.cell(row=cur_r, column=10, value=step[4])
        p_cell.border = styles["cell_border"]
        p_cell.alignment = styles["align_center"]
        if step[4] == "必須":
            p_cell.font = styles["font_fail"]
            p_cell.fill = styles["fill_fail"]
        else:
            p_cell.font = styles["font_info"]
            p_cell.fill = styles["fill_info"]

        cur_r += 1

    # Section 2: Available Features and Plugin Installation
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "2. JMeterで利用可能な主要機能と追加方法（Plugins Manager & 拡張プラグイン）", 10)
    cur_r += 1

    plugins = [
        ("Custom Thread Groups", "スレッドグループ拡張", "Plugins Manager", "Available Plugins で「Custom Thread Groups」をチェックし導入", "Stepping Thread Group や Concurrency Thread Group を提供。急激な負荷ではなく現実的な段階的負荷（ランプアップ・維持・ランプダウン）を自在に設計可能。実務必須。"),
        ("3 Basic Graphs", "基本グラフリスナー", "Plugins Manager", "Available Plugins で「3 Basic Graphs」をチェックし導入", "Active Threads Over Time (アクティブスレッド推移), Response Times Over Time (応答時間推移), Transactions per Second (TPS推移) をリアルタイム表示。"),
        ("5 Additional Graphs", "拡張グラフリスナー", "Plugins Manager", "Available Plugins で「5 Additional Graphs」をチェックし導入", "Response Codes per Second (ステータスコード別推移), Latency Over Time (ネットワークレイテンシ推移), Connect Times Over Time (TCP接続時間推移) などを可視化。"),
        ("Custom JMeter Functions", "追加組み込み関数", "Plugins Manager", "Available Plugins で「Custom JMeter Functions」をチェックし導入", "${__base64Encode()}, ${__base64Decode()}, ${__UUID()}, ${__time()} などの強力な文字列・日時操作関数を追加。トークン認証や暗号化ヘッダ生成に重宝。"),
        ("JSON/YAML Plugins", "JSON処理拡張", "Plugins Manager", "Available Plugins で「JSON/YAML Plugins」をチェックし導入", "JSON Format PostProcessor や JSON Path Assertion 等、REST API の複雑なレスポンス検証・抽出を効率化。"),
        ("Dummy Sampler", "ダミーサンプラー", "Plugins Manager", "Available Plugins で「Dummy Sampler」をチェックし導入", "実際のネットワーク通信を行わず、指定したレスポンスコード・レスポンス本文・応答時間を模擬返却するモックサンプラー。シナリオ単体検証やデバッグに極めて有効。"),
        ("JSR223 Groovy", "スクリプト拡張機能", "標準機能 (組込済)", "追加不要。Sampler / PreProcessor / PostProcessor で Language「groovy」を選択", "Javaに極めて近いGroovyスクリプトを実行可能。複雑な認証トークン計算、HMAC署名生成、動的パラメータ加工、独自アサーションを最高速で実行可能。"),
        ("Regular Expression Extractor", "正規表現抽出", "標準機能 (組込済)", "追加不要。各サンプラー配下に Add > Post Processors > Regular Expression Extractor", "レスポンス本文・ヘッダーから正規表現で特定文字列（セッショントークン、ID等）を抽出し変数に保持する標準機能。高速でオーバーヘッドが極小。")
    ]

    tbl_headers2_p = ["機能・プラグイン名", "カテゴリ", "提供形態", "追加・利用方法", "機能概要・現場での活用メリット"]
    ws.row_dimensions[cur_r].height = 24
    cols_map2_p = [(1, 2), (3, 3), (4, 4), (5, 6), (7, 10)]
    for i, (sc, ec) in enumerate(cols_map2_p):
        if sc != ec:
            ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_tbl_header"]
            cell.border = styles["header_border"]
            if c == sc:
                cell.value = tbl_headers2_p[i]
                cell.font = styles["font_tbl_header"]
                cell.alignment = styles["align_center"]

    cur_r += 1
    for idx, plg in enumerate(plugins):
        ws.row_dimensions[cur_r].height = 36
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # Name
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
        for c in range(1, 3):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = plg[0]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Category
        c_cell = ws.cell(row=cur_r, column=3, value=plg[1])
        c_cell.fill = fill_curr
        c_cell.border = styles["cell_border"]
        c_cell.font = styles["font_body"]
        c_cell.alignment = styles["align_center"]

        # Type
        c_cell = ws.cell(row=cur_r, column=4, value=plg[2])
        c_cell.fill = fill_curr
        c_cell.border = styles["cell_border"]
        c_cell.font = styles["font_body"]
        c_cell.alignment = styles["align_center"]

        # How to add
        ws.merge_cells(start_row=cur_r, start_column=5, end_row=cur_r, end_column=6)
        for c in range(5, 7):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 5:
                c_cell.value = plg[3]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        # Advantage
        ws.merge_cells(start_row=cur_r, start_column=7, end_row=cur_r, end_column=10)
        for c in range(7, 11):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 7:
                c_cell.value = plg[4]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        cur_r += 1

    # Image Area for Sheet 1
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "3. Windows 11 環境構築・プラグイン導入証跡貼付エリア", 10)
    cur_r += 1
    styles["apply_image_placeholder"](
        ws, cur_r, 1, cur_r + 8, 5,
        "Windows 環境変数設定画面 & java/jmeter コマンド実行結果",
        "システム環境変数の JAVA_HOME / PATH / JMETER_HOME 設定画面および PowerShell での java -version / jmeter -v 実行証跡キャプチャ"
    )
    styles["apply_image_placeholder"](
        ws, cur_r, 6, cur_r + 8, 10,
        "JMeter Plugins Manager 画面 & プラグイン導入証跡",
        "Options > Plugins Manager を開き、「Installed Plugins」または「Available Plugins」でプラグインを選択・適用した状態のキャプチャ"
    )
    cur_r += 9

    # Set column widths
    col_widths2 = [8, 16, 14, 14, 16, 16, 16, 16, 16, 12]
    for idx, w in enumerate(col_widths2, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
