# -*- coding: utf-8 -*-
"""
Sheet 03: Cognito・APIGW連携 (Cognito OAuth2 & API Gateway Integration)
"""

from openpyxl.utils import get_column_letter

def build_sheet_cognito_apigw(wb, styles):
    ws = wb.create_sheet(title="03_Cognito・APIGW連携")
    ws.views.sheetView[0].showGridLines = True

    styles["apply_title_banner"](
        ws,
        "03. AWS Cognito認証＆API Gateway連携シナリオ構築実践ガイド",
        "OAuth 2.0 Client Credentials・BASE64ヘッダ生成・正規表現抽出・X-API-ID付与・後続API連携",
        10
    )

    cur_r = 4
    styles["apply_section_header"](ws, cur_r, "1. AWS Cognito OAuth 2.0 認証＆API Gateway 連携の全体フロー", 10)
    cur_r += 1

    flow_steps = [
        ("Step 1", "変数定義 (UDV)", "JMeter内部", "テスト計画配下の「User Defined Variables」にドメイン、ClientID、ClientSecret、X-API-ID等の接続情報を一括定義。"),
        ("Step 2", "Basic認証ヘッダ生成", "JMeter内部", "JMeter標準関数 ${__base64Encode(${CLIENT_ID}:${CLIENT_SECRET})} を用い、Authorizationヘッダ文字列を動的生成。"),
        ("Step 3", "Cognito Token要求", "JMeter → Cognito", "POST /oauth2/token に grant_type=client_credentials、Content-Type: application/x-www-form-urlencoded で送信。"),
        ("Step 4", "Token抽出・保持", "JMeter内部", "「Regular Expression Extractor (正規表現抽出)」を用い、レスポンスBodyから access_token を抽出し変数に保持。"),
        ("Step 5", "API Gateway呼出", "JMeter → APIGW", "HTTP Header Managerで「X-API-ID: ${X_API_ID}」「Authorization: Bearer ${ACCESS_TOKEN}」を付与し業務APIへ送信。"),
        ("Step 6", "応答検証 (Assertion)", "JMeter内部", "Response Assertion で HTTP 200 OK およびレスポンス本文の業務データ存在を自動検証。")
    ]

    for st, title, comm, desc in flow_steps:
        ws.row_dimensions[cur_r].height = 28
        ws.cell(row=cur_r, column=1, value=st).font = styles["font_body_bold"]
        ws.cell(row=cur_r, column=1).fill = styles["fill_card"]
        ws.cell(row=cur_r, column=1).border = styles["cell_border"]
        ws.cell(row=cur_r, column=1).alignment = styles["align_center"]

        ws.cell(row=cur_r, column=2, value=title).font = styles["font_body_bold"]
        ws.cell(row=cur_r, column=2).fill = styles["fill_card"]
        ws.cell(row=cur_r, column=2).border = styles["cell_border"]
        ws.cell(row=cur_r, column=2).alignment = styles["align_left"]

        ws.merge_cells(start_row=cur_r, start_column=3, end_row=cur_r, end_column=4)
        for c in range(3, 5):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_card"]
            c_cell.border = styles["cell_border"]
            if c == 3:
                c_cell.value = comm
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_center"]

        ws.merge_cells(start_row=cur_r, start_column=5, end_row=cur_r, end_column=10)
        for c in range(5, 11):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_white"]
            c_cell.border = styles["cell_border"]
            if c == 5:
                c_cell.value = desc
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]
        cur_r += 1

    # Section 2: Detailed Component Configuration Table
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "2. JMeter各コンポーネントの設定パラメータ詳細一覧（完全設定マニュアル）", 10)
    cur_r += 1

    configs = [
        ("User Defined Variables", "テスト計画直下", "テスト環境変数の一括定義", 
         "・COGNITO_DOMAIN: auth-stg.example.amazoncognito.com\n"
         "・CLIENT_ID: 1example23456789abcdef\n"
         "・CLIENT_SECRET: 1234567890abcdefghijklmnopqrstuvwxyz\n"
         "・APIGW_HOST: api-stg.example.com\n"
         "・X_API_ID: my-api-key-998877\n"
         "※後から環境が変わってもこの1箇所の変更で全リクエストに反映可能。", "共通設定"),
        ("HTTP Header Manager (Cognito用)", "Cognitoサンプラー直下", "OAuth2トークン発行用ヘッダー",
         "【ヘッダ設定値】\n"
         "・Content-Type: application/x-www-form-urlencoded\n"
         "・Authorization: Basic ${__base64Encode(${CLIENT_ID}:${CLIENT_SECRET})}\n"
         "※JMeter 5.x標準関数 __base64Encode により、CLIENT_ID と CLIENT_SECRET をコロンで連結した文字列を動的にBase64エンコードしてBasic認証ヘッダを生成します。", "必須"),
        ("HTTP Request (01_Auth_CognitoToken)", "スレッドグループ直下", "Cognitoトークンエンドポイント呼出",
         "【サンプラー設定値】\n"
         "・Protocol: https\n"
         "・Server Name or IP: ${COGNITO_DOMAIN}\n"
         "・Port: 443\n"
         "・Method: POST\n"
         "・Path: /oauth2/token\n"
         "【Parameters タブ】(または Body Data)\n"
         "・Name: grant_type, Value: client_credentials\n"
         "・Name: scope, Value: resource-server/read (スコープ指定が必要な場合)", "必須"),
        ("Regular Expression Extractor", "01_Auth_CognitoToken 直下 (Post Processors)", "正規表現によるトークン抽出と保持",
         "【正規表現抽出器の設定値】\n"
         "・Apply to: Main sample only\n"
         "・Field to check: Body\n"
         "・Name of created variable: ACCESS_TOKEN\n"
         "・Regular Expression: \"access_token\"\\s*:\\s*\"([^\"]+)\"\n"
         "・Template: $1$\n"
         "・Match No. (0 for Random): 1\n"
         "・Default Value: TOKEN_EXTRACT_ERROR\n"
         "※抽出されたトークンは ${ACCESS_TOKEN} として後続の全サンプラーで利用可能。", "極めて重要"),
        ("(参考) JSON Extractor", "01_Auth_CognitoToken 直下 (Post Processors)", "JSONPathによるトークン抽出 (代替手段)",
         "【JSON抽出器の設定値】(正規表現抽出の代替えとして利用可能)\n"
         "・Names of created variables: JSON_TOKEN\n"
         "・JSON Path expressions: $.access_token\n"
         "・Match Numbers: 1\n"
         "・Default Values: NOT_FOUND\n"
         "※JSON構造が明確な場合はJSON Extractorも可。正規表現抽出はより軽量で高速に動作します。", "推奨代替"),
        ("HTTP Header Manager (APIGW用)", "02_APIGW_Resource 直下", "API Gateway & 認証ヘッダー付与",
         "【ヘッダ設定値】\n"
         "・X-API-ID: ${X_API_ID}  (API Gatewayで要求されるAPI識別子・APIキー)\n"
         "・Authorization: Bearer ${ACCESS_TOKEN}  (Cognitoで取得したアクセストークン)\n"
         "・Content-Type: application/json\n"
         "・Accept: application/json", "必須"),
        ("HTTP Request (02_APIGW_Resource)", "スレッドグループ直下", "API Gateway 業務リソース呼出",
         "【サンプラー設定値】\n"
         "・Protocol: https\n"
         "・Server Name or IP: ${APIGW_HOST}\n"
         "・Port: 443\n"
         "・Method: GET (または POST)\n"
         "・Path: /v1/customers/profile\n"
         "・(POSTの場合) Body Data: {\"query\": \"sample\"}", "必須"),
        ("Response Assertion", "02_APIGW_Resource 直下 (Assertions)", "正常応答とレスポンス検証",
         "【アサーション設定値】\n"
         "・Field to Test: Response Code, Patterns: 200 (HTTP正常)\n"
         "・Field to Test: Text Response, Patterns: \"customerId\": (業務データ存在確認)\n"
         "※単に200が返るだけでなく、意図通りの顧客JSONが返却されているかを厳密判定。", "必須")
    ]

    tbl_headers3_c = ["対象コンポーネント", "配置場所 (階層)", "役割・設定意図", "詳細設定プロパティ・記述内容", "種別"]
    ws.row_dimensions[cur_r].height = 24
    cols_map3_c = [(1, 2), (3, 3), (4, 5), (6, 9), (10, 10)]
    for i, (sc, ec) in enumerate(cols_map3_c):
        if sc != ec:
            ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_tbl_header"]
            cell.border = styles["header_border"]
            if c == sc:
                cell.value = tbl_headers3_c[i]
                cell.font = styles["font_tbl_header"]
                cell.alignment = styles["align_center"]

    cur_r += 1
    for idx, cfg in enumerate(configs):
        ws.row_dimensions[cur_r].height = 68
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # Component
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
        for c in range(1, 3):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = cfg[0]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Location
        c_cell = ws.cell(row=cur_r, column=3, value=cfg[1])
        c_cell.fill = fill_curr
        c_cell.border = styles["cell_border"]
        c_cell.font = styles["font_body"]
        c_cell.alignment = styles["align_center"]

        # Role
        ws.merge_cells(start_row=cur_r, start_column=4, end_row=cur_r, end_column=5)
        for c in range(4, 6):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 4:
                c_cell.value = cfg[2]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Config content
        ws.merge_cells(start_row=cur_r, start_column=6, end_row=cur_r, end_column=9)
        for c in range(6, 10):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 6:
                c_cell.value = cfg[3]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        # Priority
        p_cell = ws.cell(row=cur_r, column=10, value=cfg[4])
        p_cell.border = styles["cell_border"]
        p_cell.alignment = styles["align_center"]
        if cfg[4] in ["必須", "極めて重要"]:
            p_cell.font = styles["font_fail"]
            p_cell.fill = styles["fill_fail"]
        else:
            p_cell.font = styles["font_info"]
            p_cell.fill = styles["fill_info"]

        cur_r += 1

    # Section 3: Regular Expression Extractor Deep Dive
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "3. 正規表現抽出 (Regular Expression Extractor) の設定値完全解説", 10)
    cur_r += 1

    regex_params = [
        ("参照名 (Reference Name)", "ACCESS_TOKEN", "抽出した値を格納するJMeter変数名。後続リクエストで ${ACCESS_TOKEN} として参照可能になります。空欄だとエラーになります。"),
        ("正規表現 (Regular Expression)", "\"access_token\"\\s*:\\s*\"([^\"]+)\"", 
         "・\"access_token\": JSON内のキー名。\n"
         "・\\s*:\\s*: コロンの前後に半角空白や改行が何個あっても許容する指定。\n"
         "・\"([^\"]+)\": ダブルクォーテーションで囲まれた「閉じクォート以外の任意の1文字以上」をグループ1としてキャプチャ。トークン文字列本体を正確に切り出します。"),
        ("テンプレート (Template)", "$1$", 
         "正規表現内のどのキャプチャグループ（括弧）を変数に代入するかを指定します。\n"
         "・$1$: 1番目の括弧 ([^\"]+) の内容のみを抽出 (トークン文字列そのもの)。\n"
         "※$0$ を指定するとマッチ全体 (\"access_token\":\"eyJ...\") が格納されてしまい認証エラーになるので注意。"),
        ("一致番号 (Match No.)", "1", 
         "レスポンス内にマッチが複数あった場合の選択基準。\n"
         "・1: 最初に見つかったマッチを採用 (通常は1を指定)。\n"
         "・0: 見つかったマッチの中からランダムに選択。\n"
         "・-1: すべてのマッチを配列変数 (ACCESS_TOKEN_1, ACCESS_TOKEN_2...) として全件保持。"),
        ("デフォルト値 (Default Value)", "TOKEN_EXTRACT_ERROR", 
         "正規表現がマッチせずトークン抽出に失敗した際に変数へセットされるフォールバック値。\n"
         "・設定の意義: 空白にしておくと変数自体が未定義 (${ACCESS_TOKEN} のまま) となりデバッグが困難になります。明示的なエラー文字列を入れておくことで、後続APIが 401 Unauthorized となり、即座に抽出失敗と特定できます。")
    ]

    tbl_headers3_r = ["設定プロパティ名", "推奨設定値", "初心者向け技術解説・実務上の注意点"]
    ws.row_dimensions[cur_r].height = 24
    cols_map3_r = [(1, 3), (4, 5), (6, 10)]
    for i, (sc, ec) in enumerate(cols_map3_r):
        if sc != ec:
            ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_tbl_header"]
            cell.border = styles["header_border"]
            if c == sc:
                cell.value = tbl_headers3_r[i]
                cell.font = styles["font_tbl_header"]
                cell.alignment = styles["align_center"]

    cur_r += 1
    for idx, rp in enumerate(regex_params):
        ws.row_dimensions[cur_r].height = 42
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # Name
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=3)
        for c in range(1, 4):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = rp[0]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Value
        ws.merge_cells(start_row=cur_r, start_column=4, end_row=cur_r, end_column=5)
        for c in range(4, 6):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 4:
                c_cell.value = rp[1]
                c_cell.font = styles["font_code"]
                c_cell.alignment = styles["align_center"]

        # Explanation
        ws.merge_cells(start_row=cur_r, start_column=6, end_row=cur_r, end_column=10)
        for c in range(6, 11):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 6:
                c_cell.value = rp[2]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        cur_r += 1

    # Section 4: BASE64 Encoding Methods Comparison
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "4. Client ID & Secret の BASE64 エンコード手法の比較", 10)
    cur_r += 1

    base64_methods = [
        ("手法 A: JMeter標準組み込み関数 (最推奨)", "Authorization: Basic ${__base64Encode(${CLIENT_ID}:${CLIENT_SECRET})}", 
         "HTTP Header Manager のヘッダー値に上記関数を直接記述する。\n"
         "・メリット: プラグイン導入不要、JMeter 5.x標準機能、設定が1行で完結しオーバーヘッドが極小。\n"
         "・注意点: 変数名 (${CLIENT_ID}) のコロン前後に余計な半角空白を入れないこと。", "◎ 最も推奨"),
        ("手法 B: JSR223 PreProcessor (Groovy)", 
         "// サンプラー直下の JSR223 PreProcessor に記述\n"
         "String raw = vars.get('CLIENT_ID') + ':' + vars.get('CLIENT_SECRET');\n"
         "String encoded = raw.bytes.encodeBase64().toString();\n"
         "vars.put('BASIC_AUTH_HEADER', 'Basic ' + encoded);\n"
         "// Header Manager で Authorization: ${BASIC_AUTH_HEADER} と参照",
         "・メリット: 複雑なOAuth署名生成や動的ソルト付与、条件分岐処理にも柔軟に対応可能。\n"
         "・注意点: Groovyスクリプトの基礎知識が必要。", "○ 応用向け"),
        ("手法 C: 外部ツールでの事前エンコード", "User Defined Variables で BASIC_AUTH_HEADER に生成済みの Base64 文字列を直書き", 
         "・メリット: 実行時に関数計算が発生しないため極めて軽量。\n"
         "・注意点: クライアント認証情報が固定値となり、CSVによる動的マルチアカウント認証には利用できない。", "△ 固定アカウント専用")
    ]

    for bm in base64_methods:
        ws.row_dimensions[cur_r].height = 48
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=3)
        for c in range(1, 4):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_card"]
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = bm[0]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        ws.merge_cells(start_row=cur_r, start_column=4, end_row=cur_r, end_column=6)
        for c in range(4, 7):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_code"]
            c_cell.border = styles["cell_border"]
            if c == 4:
                c_cell.value = bm[1]
                c_cell.font = styles["font_code"]
                c_cell.alignment = styles["align_left"]

        ws.merge_cells(start_row=cur_r, start_column=7, end_row=cur_r, end_column=9)
        for c in range(7, 10):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_white"]
            c_cell.border = styles["cell_border"]
            if c == 7:
                c_cell.value = bm[2]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        p_cell = ws.cell(row=cur_r, column=10, value=bm[3])
        p_cell.border = styles["cell_border"]
        p_cell.alignment = styles["align_center"]
        p_cell.font = styles["font_body_bold"]
        p_cell.fill = styles["fill_pass"] if "最推奨" in bm[3] else styles["fill_sub"]
        cur_r += 1

    # Section 5: Troubleshooting & Verification
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "5. 単体デバッグ検証手順とよくあるエラーの対処チェックリスト", 10)
    cur_r += 1

    troubles = [
        ("Cognito呼出で HTTP 400 Bad Request", "リクエスト形式またはBodyパラメータの誤り", 
         "・HTTP Header Manager で Content-Type: application/x-www-form-urlencoded が設定されているか？\n"
         "・grant_type=client_credentials が「Parameters」タブまたは正しくURLエンコードされたBodyに記述されているか？"),
        ("Cognito呼出で HTTP 401 Unauthorized", "Basic認証ヘッダのBase64エンコード不整合", 
         "・Authorization ヘッダが「Basic <Base64文字列>」の形式になっているか (Basicと文字列の間に半角スペース必須)。\n"
         "・Client ID と Client Secret を連結するコロン (:) の前後に余計な空白文字が入っていないか確認。"),
        ("API Gateway呼出で HTTP 403 Forbidden", "X-API-ID の不一致またはスコープ不足", 
         "・HTTP Header Manager で「X-API-ID: ${X_API_ID}」が正しくサンプラーに付与されているか？ (キー名の大文字小文字に注意)\n"
         "・Cognito オーソライザーのリソースサーバー設定 (カスタムスコープ) がAPI Gatewayのメソッド要求と一致しているか確認。"),
        ("後続リクエストで Authorization が空文字", "正規表現抽出の失敗またはスコープ範囲の誤り", 
         "・View Results Tree で 01_Auth_CognitoToken のレスポンスを開き、\"access_token\" キーが実際に含まれているか確認。\n"
         "・正規表現抽出器の Apply to が「Main sample only」になっているか確認 (Sub-samples onlyになっていると抽出不能)。")
    ]

    for tb in troubles:
        ws.row_dimensions[cur_r].height = 42
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=3)
        for c in range(1, 4):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_card"]
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = tb[0]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        ws.merge_cells(start_row=cur_r, start_column=4, end_row=cur_r, end_column=5)
        for c in range(4, 6):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_white"]
            c_cell.border = styles["cell_border"]
            if c == 4:
                c_cell.value = tb[1]
                c_cell.font = styles["font_warn"]
                c_cell.alignment = styles["align_left"]

        ws.merge_cells(start_row=cur_r, start_column=6, end_row=cur_r, end_column=10)
        for c in range(6, 11):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_white"]
            c_cell.border = styles["cell_border"]
            if c == 6:
                c_cell.value = tb[2]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]
        cur_r += 1

    # Image Area for Sheet 3
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "6. Cognito認証＆API Gateway連携デバッグ証跡貼付エリア", 10)
    cur_r += 1
    styles["apply_image_placeholder"](
        ws, cur_r, 1, cur_r + 8, 5,
        "Cognito トークン取得設定＆正規表現抽出 (Regular Expression Extractor) 画面",
        "Cognito HTTP Requestサンプラーの設定、HTTP Header Manager (Basic認証関数)、および正規表現抽出器の設定画面キャプチャ"
    )
    styles["apply_image_placeholder"](
        ws, cur_r, 6, cur_r + 8, 10,
        "API Gateway 呼出ヘッダ (X-API-ID & Bearer Token)＆View Results Tree疎通証跡",
        "View Results Tree の [Request] タブで X-API-ID と Bearer ${ACCESS_TOKEN} が正しく展開され、HTTP 200 OK で疎通した証跡キャプチャ"
    )
    cur_r += 9

    # Set column widths
    col_widths3 = [8, 16, 14, 14, 16, 16, 16, 16, 16, 12]
    for idx, w in enumerate(col_widths3, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
