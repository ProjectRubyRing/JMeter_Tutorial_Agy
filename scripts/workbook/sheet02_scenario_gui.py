# -*- coding: utf-8 -*-
"""
Sheet 02: シナリオ作成手法 (GUI Scenario Creation & Debugging)
"""

from openpyxl.utils import get_column_letter

def build_sheet_scenario_gui(wb, styles):
    ws = wb.create_sheet(title="02_シナリオ作成手法")
    ws.views.sheetView[0].showGridLines = True

    styles["apply_title_banner"](
        ws,
        "02. JMeter GUI シナリオ作成・画面実行・結果確認・注意点ガイド",
        "テスト計画構造・主要コンポーネント・画面実行手順・リスナー結果確認・確認観点・GUI実行禁止則",
        10
    )

    cur_r = 4
    styles["apply_section_header"](ws, cur_r, "1. JMeter テスト計画の階層構造と主要コンポーネント解説", 10)
    cur_r += 1

    components = [
        ("Test Plan", "テスト計画", "最上位のルート要素。シナリオ全体の共通設定、ユーザー定義変数 (User Defined Variables)、テスト実行順序などを統括する。", "右クリック > Add > Config Element 等で全体変数を定義"),
        ("Thread Group", "スレッドグループ", "仮想ユーザーの集合。同時実行ユーザー数 (Number of Threads)、ランプアップ時間 (Ramp-up Period)、ループ回数 (Loop Count)、測定継続時間 (Duration) を定義する。", "負荷テストの負荷プロファイルを決定する中核要素"),
        ("HTTP Request Defaults", "HTTP初期値設定", "全HTTPサンプラーで共通するホスト名/IP、ポート番号、プロトコル (https/http)、エンコーディング (UTF-8) を一括定義する。", "ホストが変わっても1箇所の変更で全リクエストに反映可能"),
        ("HTTP Header Manager", "HTTPヘッダマネージャ", "リクエストヘッダーを一括付与する。「Content-Type: application/json」や「Authorization: Bearer ${TOKEN}」等を定義。", "APIテストやREST通信では必須となるコンポーネント"),
        ("HTTP Cookie Manager", "HTTPクッキーマネージャ", "サーバから発行されたCookie (JSESSIONID, セッションCookie等) を自動保持し、後続リクエストへ自動送信する。", "Web画面遷移やログイン状態を維持する場合に必須"),
        ("CSV Data Set Config", "CSVデータセット設定", "外部CSVファイルを読み込み、行ごとに異なるデータ (ユーザーID, パスワード, 検索キーワード等) を変数として割り当てる。", "マルチユーザーによる負荷テストで同一データ衝突を防ぐ必須機能"),
        ("HTTP Request", "HTTPサンプラー", "実際に送信する個々のHTTP/HTTPSリクエスト。パス (/api/v1/login)、メソッド (GET/POST/PUT)、Bodyデータ (JSON) などを指定する。", "測定対象となる各画面やAPIごとに1つずつ作成する"),
        ("Response Assertion", "レスポンスアサーション", "サーバからの応答が正常であるかを自動検証する。HTTPステータスコードが200か、本文に特定文字列が含まれるかを判定。", "単に通信できただけでなく、期待通りの画面/JSONが返ったかを判定"),
        ("Uniform Random Timer", "均等ランダムタイマー", "リクエスト間に人間らしい「思考時間 (Think Time)」を挿入する。(例: 最小1秒〜最大3秒のランダム待機)。", "タイマーがないと非現実的なDDOS攻撃となり正確な測定が不能"),
        ("View Results Tree", "結果をツリーで表示", "デバッグ用リスナー。リクエストヘッダ・Body、レスポンスヘッダ・Body・ステータスコードを完全に可視化する。", "※注意: 単体デバッグ時のみ使用。高負荷テスト時は必ず無効化")
    ]

    tbl_headers3 = ["英語名", "日本語名", "機能解説・役割", "実務上の活用ポイント・注意点"]
    ws.row_dimensions[cur_r].height = 24
    cols_map3 = [(1, 2), (3, 4), (5, 7), (8, 10)]
    for i, (sc, ec) in enumerate(cols_map3):
        if sc != ec:
            ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_tbl_header"]
            cell.border = styles["header_border"]
            if c == sc:
                cell.value = tbl_headers3[i]
                cell.font = styles["font_tbl_header"]
                cell.alignment = styles["align_center"]

    cur_r += 1
    for idx, comp in enumerate(components):
        ws.row_dimensions[cur_r].height = 36
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # Eng Name
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
        for c in range(1, 3):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = comp[0]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Jp Name
        ws.merge_cells(start_row=cur_r, start_column=3, end_row=cur_r, end_column=4)
        for c in range(3, 5):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 3:
                c_cell.value = comp[1]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Description
        ws.merge_cells(start_row=cur_r, start_column=5, end_row=cur_r, end_column=7)
        for c in range(5, 8):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 5:
                c_cell.value = comp[2]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        # Practical Points
        ws.merge_cells(start_row=cur_r, start_column=8, end_row=cur_r, end_column=10)
        for c in range(8, 11):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 8:
                c_cell.value = comp[3]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]
        cur_r += 1

    # Section 2: Walkthrough & GUI Debug Procedure
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "2. シナリオ作成手順とGUI画面での実行・単体デバッグ操作", 10)
    cur_r += 1

    gui_steps = [
        ("手順 1", "テスト計画の新規作成 & 初期値定義",
         "1. JMeter起動後、「Test Plan」を選択し名前をわかりやすく変更 (例: TP_EC_UserScenario)。\n"
         "2. Test Planを右クリック > Add > Config Element > HTTP Request Defaults を追加。\n"
         "   - Server Name or IP: 接続先ドメインまたはIP (例: api-stg.example.local)\n"
         "   - Port Number: 443 (HTTPS) または 80 (HTTP)\n"
         "   - Protocol: https\n"
         "   - Content encoding: UTF-8\n"
         "3. Test Planを右クリック > Add > Config Element > HTTP Header Manager を追加。\n"
         "   - Name: Content-Type, Value: application/json\n"
         "4. Cookieが必要な場合は HTTP Cookie Manager を追加 (Clear cookies each iteration にチェック)。"),
        ("手順 2", "スレッドグループの追加と単体デバッグ設定",
         "1. Test Planを右クリック > Add > Threads (Users) > Thread Group を追加。\n"
         "2. 本番負荷ではなく「単体デバッグ用」として以下を設定:\n"
         "   - Number of Threads (users): 1 (必ず1ユーザーで開始)\n"
         "   - Ramp-up period (seconds): 0\n"
         "   - Loop Count: 1\n"
         "※本番用の負荷パラメータ (スレッド数50, ランプアップ60秒, 時間指定1200秒など) はプロパティ化 (${__P(threads, 50)}) してCLI引数で制御します。"),
        ("手順 3", "HTTPサンプラーの作成とアサーション追加",
         "1. Thread Groupを右クリック > Add > Sampler > HTTP Request を追加。\n"
         "   - サンプラー名: 01_Login_API (ナンバリングすると分析レポートで識別しやすい)\n"
         "   - Method: POST, Path: /api/v1/auth/login\n"
         "   - Body Data にJSONペイロードを入力 (例: {\"username\":\"${USER_ID}\", \"password\":\"${PASSWORD}\"})\n"
         "2. HTTP Requestを右クリック > Add > Assertions > Response Assertion を追加。\n"
         "   - Field to Test: Response Code, Patterns to Test: 200 を入力。\n"
         "3. 業務エラーを検知するため、もう1つ Response Assertion を追加:\n"
         "   - Field to Test: Text Response, Patterns to Test: \"success\": true などの必須文字列を指定。"),
        ("手順 4", "CSVデータセット設定による複数ユーザーデータ読み込み",
         "1. Thread Groupを右クリック > Add > Config Element > CSV Data Set Config を追加。\n"
         "   - Filename: 相対パス (例: ../csv/users.csv) または絶対パス\n"
         "   - File encoding: UTF-8\n"
         "   - Variable Names: USER_ID,PASSWORD,EXPECTED_ROLE (カンマ区切り)\n"
         "   - Delimiter: ,\n"
         "   - Recycle on EOF?: True (末尾に達したら先頭に戻る)\n"
         "   - Stop thread on EOF?: False"),
        ("手順 5", "画面での実行操作（ツールバー・ログビューア・結果クリア）",
         "1. 実行方法:\n"
         "   - 上部ツールバーの緑色の「Start」ボタン (再生アイコン)、またはショートカット [Ctrl + R] で実行。\n"
         "   - エラーを無視して強制実行する場合は「Start no pauses」[Ctrl + Shift + R]。\n"
         "2. 停止方法:\n"
         "   - 赤色の「Stop」ボタン: 即座に通信切断 (非推奨)。\n"
         "   - 黄色の「Shutdown」ボタン: 現在のサンプラー完了を待って安全停止 (推奨)。\n"
         "3. ログビューアの確認:\n"
         "   - 画面右上にある「黄色い三角形 (感嘆符)」アイコンをクリックして下部にログパネルを表示。\n"
         "   - スクリプト構文エラーやファイルが見つからない例外が発生していないかリアルタイム確認。\n"
         "4. 実行結果のクリア:\n"
         "   - 「ほうき1本」アイコン: 選択中リスナーの結果をクリア。\n"
         "   - 「ほうき2本」アイコン [Ctrl + Shift + E]: 全リスナーの結果を全消去。再実行前に必ず押下。"),
        ("手順 6", "シナリオ (.jmx) の出力・保存",
         "1. 上部メニュー「File」>「Save Test Plan as...」を選択。\n"
         "2. わかりやすいファイル名 (例: sample_scenario.jmx) で作業フォルダに保存。\n"
         "3. 【絶対厳守】Linux CLI実行前に、View Results Tree や Summary Report などのリスナーを右クリック >「Disable (無効化)」するか、JMXから削除します。(GUI描画負荷によるメモリ枯渇を完全排除するため)。")
    ]

    for step in gui_steps:
        ws.row_dimensions[cur_r].height = 54
        ws.cell(row=cur_r, column=1, value=step[0]).font = styles["font_body_bold"]
        ws.cell(row=cur_r, column=1).fill = styles["fill_card"]
        ws.cell(row=cur_r, column=1).border = styles["cell_border"]
        ws.cell(row=cur_r, column=1).alignment = styles["align_center"]

        ws.cell(row=cur_r, column=2, value=step[1]).font = styles["font_body_bold"]
        ws.cell(row=cur_r, column=2).fill = styles["fill_card"]
        ws.cell(row=cur_r, column=2).border = styles["cell_border"]
        ws.cell(row=cur_r, column=2).alignment = styles["align_left"]

        ws.merge_cells(start_row=cur_r, start_column=3, end_row=cur_r, end_column=10)
        for c in range(3, 11):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_white"]
            c_cell.border = styles["cell_border"]
            if c == 3:
                c_cell.value = step[2]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]
        cur_r += 1

    # Section 3: Listener Result Verification & Pitfalls
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "3. 結果レポート（リスナー）の確認方法・確認観点・重大注意点", 10)
    cur_r += 1

    listener_guides = [
        ("View Results Tree (結果をツリーで表示)", "単体デバッグ", "・[Sampler result] タブ: HTTPステータス (200 OK)、レスポンスタイム (Load time)、データサイズを確認。\n・[Request] タブ: 実際に送信されたURL、ヘッダ、Body内の変数 (${USER_ID}) が正しく置換されているか確認。\n・[Response data] タブ: サーバから意図したJSON/HTMLが返っているか確認。\n・画面左下のドロップダウンを「JSON Path Tester」や「RegExp Tester」に切り替えて抽出式のテストが可能。", "単体デバッグ専用。全リクエストのBodyをオンメモリ保持するため、高負荷実行時に有効化すると数分でOut Of Memoryで強制停止する。本番実行時は必ず無効化。"),
        ("Summary Report / Aggregate Report", "統計サマリ確認", "・Samples: リクエスト総試行回数。\n・Average: 平均応答時間 (ミリ秒)。\n・Min / Max: 最小および最大応答時間。\n・90% / 95% / 99% Line (Percentile): 全体の90%/95%/99%のユーザーがこの時間内に応答を得た境界値。\n・Error %: アサーション失敗またはHTTPエラーの発生率。\n・Throughput: 1秒あたりの処理完了数 (TPS)。", "小規模な検証で統計値を確認するのに便利だが、高負荷テスト時はGUI集計の描画処理がCPUを消費するため、本番負荷はCLI (-l result.jtl) で出力させ、後からExcelやHTMLレポートで集計する。"),
        ("【絶対厳守】GUI実行禁止則", "高負荷実行の鉄則", "JMeter GUIは「シナリオ作成」「単体デバッグ」を行うための画面であり、「負荷テスト本番を実行するための画面」ではない。\nGUIで数千〜数万リクエストを流すと、Java Swingの画面描画スレッドがCPUを100%食いつぶし、JMeter自体がボトルネック化してレスポンス遅延や偽りのタイムアウトが多発する。", "高負荷テストは必ずLinuxサーバ上のCLI (Non-GUIモード: jmeter -n -t ... -l ...) で実行すること。GUIは開発端末でのみ利用する。")
    ]

    tbl_headers3_l = ["リスナー名 / 項目", "主用途", "確認方法・画面の見方・確認観点", "現場での注意点・重大リスク"]
    ws.row_dimensions[cur_r].height = 24
    cols_map3_l = [(1, 2), (3, 3), (4, 7), (8, 10)]
    for i, (sc, ec) in enumerate(cols_map3_l):
        if sc != ec:
            ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_tbl_header"]
            cell.border = styles["header_border"]
            if c == sc:
                cell.value = tbl_headers3_l[i]
                cell.font = styles["font_tbl_header"]
                cell.alignment = styles["align_center"]

    cur_r += 1
    for idx, lg in enumerate(listener_guides):
        ws.row_dimensions[cur_r].height = 48
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # Name
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
        for c in range(1, 3):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = lg[0]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Purpose
        c_cell = ws.cell(row=cur_r, column=3, value=lg[1])
        c_cell.fill = fill_curr
        c_cell.border = styles["cell_border"]
        c_cell.font = styles["font_body_bold"]
        c_cell.alignment = styles["align_center"]

        # How to check
        ws.merge_cells(start_row=cur_r, start_column=4, end_row=cur_r, end_column=7)
        for c in range(4, 8):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 4:
                c_cell.value = lg[2]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        # Warnings
        ws.merge_cells(start_row=cur_r, start_column=8, end_row=cur_r, end_column=10)
        for c in range(8, 11):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 8:
                c_cell.value = lg[3]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        cur_r += 1

    # Image Area for Sheet 2
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "4. シナリオ作成・デバッグ証跡貼付エリア", 10)
    cur_r += 1
    styles["apply_image_placeholder"](
        ws, cur_r, 1, cur_r + 8, 5,
        "スレッドグループ＆HTTPサンプラー設定画面",
        "スレッド数・ループ回数の設定およびHTTP RequestのMethod/Path/Body設定画面のキャプチャ"
    )
    styles["apply_image_placeholder"](
        ws, cur_r, 6, cur_r + 8, 10,
        "「結果をツリーで表示」デバッグ成功画面",
        "View Results Treeで全リクエストが緑色アイコンになり、Request/Responseが正常に疎通している証跡キャプチャ"
    )
    cur_r += 9

    # Set column widths
    col_widths3 = [10, 16, 14, 14, 16, 16, 16, 16, 16, 12]
    for idx, w in enumerate(col_widths3, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
