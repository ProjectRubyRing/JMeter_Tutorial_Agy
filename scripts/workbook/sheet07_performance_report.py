# -*- coding: utf-8 -*-
"""
Sheet 07: 性能分析レポート (Performance Analysis Report)
"""

from openpyxl.styles import Border
from openpyxl.utils import get_column_letter

def build_sheet_performance_report(wb, styles):
    ws = wb.create_sheet(title="07_性能分析レポート")
    ws.views.sheetView[0].showGridLines = True

    styles["apply_title_banner"](
        ws,
        "07. 負荷テスト性能分析レポート & SLA判定書（成果物報告テンプレート）",
        "総合KPI・SLA達成状況・トランザクション別性能統計・インフラ相関・エラー分析・改善提言",
        12
    )

    cur_r = 4
    styles["apply_section_header"](ws, cur_r, "1. 負荷テスト実施要領・前提条件サマリ", 12)
    cur_r += 1

    summary_items = [
        ("対象システム名", "Cognito OAuth2 / API Gateway / マイクロサービス基盤", "テスト実施日時", "2026年10月04日 14:00 ～ 14:20 JST (20分間)"),
        ("テストシナリオ名", "Cognito認証＆REST API業務リクエスト連続負荷テスト", "テスト実行環境", "ステージング環境 (本番同等構成)"),
        ("最大同時仮想ユーザー数", "50 スレッド (Threads)", "ランプアップ時間", "60 秒 (段階的負荷上昇)"),
        ("測定継続時間", "1,200 秒 (20 分間定常負荷)", "テストデータ規模", "CSV 1,000 件ユニークアカウント"),
        ("テスト実施者", "性能テスト推進チーム (Performance Eng Group)", "負荷生成サーバ", "Linux (AlmaLinux 9, 4 vCPU, 8GB RAM)")
    ]

    for item in summary_items:
        ws.row_dimensions[cur_r].height = 20
        # Col 1-2
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
        for c in range(1, 3):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_card"]
            cell.border = styles["cell_border"]
            if c == 1:
                cell.value = item[0]
                cell.font = styles["font_body_bold"]
                cell.alignment = styles["align_left"]
        # Col 3-6
        ws.merge_cells(start_row=cur_r, start_column=3, end_row=cur_r, end_column=6)
        for c in range(3, 7):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_white"]
            cell.border = styles["cell_border"]
            if c == 3:
                cell.value = item[1]
                cell.font = styles["font_body"]
                cell.alignment = styles["align_left"]
        # Col 7-8
        ws.merge_cells(start_row=cur_r, start_column=7, end_row=cur_r, end_column=8)
        for c in range(7, 9):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_card"]
            cell.border = styles["cell_border"]
            if c == 7:
                cell.value = item[2]
                cell.font = styles["font_body_bold"]
                cell.alignment = styles["align_left"]
        # Col 9-12
        ws.merge_cells(start_row=cur_r, start_column=9, end_row=cur_r, end_column=12)
        for c in range(9, 13):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_white"]
            cell.border = styles["cell_border"]
            if c == 9:
                cell.value = item[3]
                cell.font = styles["font_body"]
                cell.alignment = styles["align_left"]
        cur_r += 1

    # Section 2: Executive Summary & KPI Cards
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "2. エグゼクティブサマリ・総合合否判定＆主要KPIカード", 12)
    cur_r += 1

    # Overall Verdict
    ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=12)
    v_cell = ws.cell(row=cur_r, column=1)
    v_cell.value = " 【総合判定】: 合格 (PASS) - 全主要業務トランザクションにおいてSLA要件（応答時間2.0秒以内・エラー率0.1%未満）を達成"
    v_cell.font = styles["font_pass"]
    v_cell.fill = styles["fill_pass"]
    v_cell.alignment = styles["align_left"]
    v_cell.border = styles["cell_border"]
    ws.row_dimensions[cur_r].height = 28
    cur_r += 2

    # KPI Cards (6 metrics)
    kpis = [
        ("総リクエスト数", "62,480", "req", "テスト期間通算の処理件数"),
        ("全体平均応答時間", "184", "ms", "全トランザクション平均"),
        ("90%ile 応答時間", "312", "ms", "9割のユーザーが享受した応答速度"),
        ("最大スループット", "64.2", "TPS", "ピーク時の秒間トランザクション"),
        ("総エラー件数", "13", "件", "エラー発生回数 (タイムアウト等)"),
        ("総合エラー率", "0.02%", "", "許容SLA (0.10%未満) を大幅達成")
    ]

    # Row 1: KPI Titles
    ws.row_dimensions[cur_r].height = 18
    ws.row_dimensions[cur_r + 1].height = 32
    ws.row_dimensions[cur_r + 2].height = 16

    col_positions = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 10), (11, 12)]
    for i, (sc, ec) in enumerate(col_positions):
        # Title
        ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_card"]
            cell.border = styles["cell_border"]
            if c == sc:
                cell.value = kpis[i][0]
                cell.font = styles["font_kpi_lbl"]
                cell.alignment = styles["align_center"]

        # Value
        ws.merge_cells(start_row=cur_r + 1, start_column=sc, end_row=cur_r + 1, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r + 1, column=c)
            cell.fill = styles["fill_white"]
            cell.border = styles["cell_border"]
            if c == sc:
                cell.value = f"{kpis[i][1]} {kpis[i][2]}".strip()
                cell.font = styles["font_kpi_val"]
                cell.alignment = styles["align_center"]

        # Note
        ws.merge_cells(start_row=cur_r + 2, start_column=sc, end_row=cur_r + 2, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r + 2, column=c)
            cell.fill = styles["fill_white"]
            cell.border = styles["cell_border"]
            if c == sc:
                cell.value = kpis[i][3]
                cell.font = styles["font_note"]
                cell.alignment = styles["align_center"]

    cur_r += 4

    # Section 3: SLA Compliance Table
    styles["apply_section_header"](ws, cur_r, "3. SLA達成状況・性能要件合否判定表", 12)
    cur_r += 1

    sla_data = [
        ("通常業務API 応答時間", "平均応答時間 1.0 秒以内", "184 ms", "達成 (合格)", "全業務APIで平均200ms未満を維持し、極めて良好。"),
        ("通常業務API 90%ile", "90%ile 応答時間 2.0 秒以内", "312 ms", "達成 (合格)", "ピーク時でも350ms以下に収まっており、スパイク遅延なし。"),
        ("重要業務 (決済・注文) 95%ile", "95%ile 応答時間 3.0 秒以内", "485 ms", "達成 (合格)", "Cognitoトークン取得および決済処理も500ms未満で安定。"),
        ("システムエラー率", "エラー率 0.10 % 未満", "0.02 % (13件)", "達成 (合格)", "一時的なネットワーク瞬断によるリトライのみで致命的障害なし。"),
        ("目標スループット", "定常 50 TPS 以上の処理能力", "平均 52.1 TPS (Max 64.2)", "達成 (合格)", "50ユーザー同時アクセス下で目標TPSを継続的に達成。")
    ]

    tbl_headers7_sla = ["評価項目・対象", "目標SLA値・合否基準", "実測値", "判定結果", "総括・分析コメント"]
    ws.row_dimensions[cur_r].height = 24
    cols_map7_sla = [(1, 3), (4, 5), (6, 7), (8, 9), (10, 12)]
    for i, (sc, ec) in enumerate(cols_map7_sla):
        if sc != ec:
            ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_tbl_header"]
            cell.border = styles["header_border"]
            if c == sc:
                cell.value = tbl_headers7_sla[i]
                cell.font = styles["font_tbl_header"]
                cell.alignment = styles["align_center"]

    cur_r += 1
    for idx, s_row in enumerate(sla_data):
        ws.row_dimensions[cur_r].height = 24
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # Item
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=3)
        for c in range(1, 4):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = s_row[0]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Target
        ws.merge_cells(start_row=cur_r, start_column=4, end_row=cur_r, end_column=5)
        for c in range(4, 6):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 4:
                c_cell.value = s_row[1]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        # Actual
        ws.merge_cells(start_row=cur_r, start_column=6, end_row=cur_r, end_column=7)
        for c in range(6, 8):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 6:
                c_cell.value = s_row[2]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_center"]

        # Verdict
        ws.merge_cells(start_row=cur_r, start_column=8, end_row=cur_r, end_column=9)
        for c in range(8, 10):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = styles["fill_pass"]
            c_cell.border = styles["cell_border"]
            if c == 8:
                c_cell.value = s_row[3]
                c_cell.font = styles["font_pass"]
                c_cell.alignment = styles["align_center"]

        # Comments
        ws.merge_cells(start_row=cur_r, start_column=10, end_row=cur_r, end_column=12)
        for c in range(10, 13):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 10:
                c_cell.value = s_row[4]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]
        cur_r += 1

    # Section 4: Detailed Transaction Statistics Table
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "4. トランザクション別詳細性能統計データ（実数式集計）", 12)
    cur_r += 1

    tx_data = [
        ("01_Auth_GetCognitoToken", 8920, 142, 45, 1250, 210, 280, 420, 0, 0.0000, 7.43, "SLA合致"),
        ("02_APIGW_CallResource", 8920, 168, 62, 1480, 245, 310, 510, 2, 0.0002, 7.43, "SLA合致"),
        ("03_商品一覧検索", 8920, 215, 80, 2150, 310, 390, 680, 3, 0.0003, 7.43, "SLA合致"),
        ("04_商品詳細・在庫照会", 8920, 135, 38, 890, 195, 240, 360, 0, 0.0000, 7.43, "SLA合致"),
        ("05_カート投入・更新", 8920, 192, 55, 1620, 280, 350, 540, 0, 0.0000, 7.43, "SLA合致"),
        ("06_注文決済確定", 8920, 310, 110, 3210, 480, 620, 980, 8, 0.0009, 7.43, "SLA合致"),
        ("07_マイページ注文履歴", 8960, 128, 42, 780, 185, 230, 320, 0, 0.0000, 7.47, "SLA合致")
    ]

    tbl_headers7_tx = ["トランザクション名 (Label)", "試行数", "平均 (ms)", "最小 (ms)", "最大 (ms)", "90%ile", "95%ile", "99%ile", "エラー数", "エラー率", "TPS", "判定"]
    ws.row_dimensions[cur_r].height = 24
    for c_idx, h_text in enumerate(tbl_headers7_tx, 1):
        cell = ws.cell(row=cur_r, column=c_idx)
        cell.value = h_text
        cell.font = styles["font_tbl_header"]
        cell.fill = styles["fill_tbl_header"]
        cell.border = styles["header_border"]
        cell.alignment = styles["align_center"]

    start_tx_r = cur_r + 1
    cur_r += 1
    for idx, tx in enumerate(tx_data):
        ws.row_dimensions[cur_r].height = 22
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # Label
        c1 = ws.cell(row=cur_r, column=1, value=tx[0])
        c1.fill = fill_curr
        c1.font = styles["font_body_bold"]
        c1.border = styles["cell_border"]
        c1.alignment = styles["align_left"]

        # Samples
        c2 = ws.cell(row=cur_r, column=2, value=tx[1])
        c2.fill = fill_curr
        c2.font = styles["font_body"]
        c2.border = styles["cell_border"]
        c2.alignment = styles["align_right"]
        c2.number_format = "#,##0"

        # Avg
        c3 = ws.cell(row=cur_r, column=3, value=tx[2])
        c3.fill = fill_curr
        c3.font = styles["font_body"]
        c3.border = styles["cell_border"]
        c3.alignment = styles["align_right"]
        c3.number_format = "#,##0"

        # Min
        c4 = ws.cell(row=cur_r, column=4, value=tx[3])
        c4.fill = fill_curr
        c4.font = styles["font_body"]
        c4.border = styles["cell_border"]
        c4.alignment = styles["align_right"]
        c4.number_format = "#,##0"

        # Max
        c5 = ws.cell(row=cur_r, column=5, value=tx[4])
        c5.fill = fill_curr
        c5.font = styles["font_body_bold"]
        c5.border = styles["cell_border"]
        c5.alignment = styles["align_right"]
        c5.number_format = "#,##0"

        # 90%ile
        c6 = ws.cell(row=cur_r, column=6, value=tx[5])
        c6.fill = fill_curr
        c6.font = styles["font_body"]
        c6.border = styles["cell_border"]
        c6.alignment = styles["align_right"]
        c6.number_format = "#,##0"

        # 95%ile
        c7 = ws.cell(row=cur_r, column=7, value=tx[6])
        c7.fill = fill_curr
        c7.font = styles["font_body"]
        c7.border = styles["cell_border"]
        c7.alignment = styles["align_right"]
        c7.number_format = "#,##0"

        # 99%ile
        c8 = ws.cell(row=cur_r, column=8, value=tx[7])
        c8.fill = fill_curr
        c8.font = styles["font_body"]
        c8.border = styles["cell_border"]
        c8.alignment = styles["align_right"]
        c8.number_format = "#,##0"

        # Errors
        c9 = ws.cell(row=cur_r, column=9, value=tx[8])
        c9.fill = fill_curr
        c9.font = styles["font_body"]
        c9.border = styles["cell_border"]
        c9.alignment = styles["align_right"]
        c9.number_format = "#,##0"

        # Error %
        c10 = ws.cell(row=cur_r, column=10, value=tx[9])
        c10.fill = fill_curr
        c10.font = styles["font_body"]
        c10.border = styles["cell_border"]
        c10.alignment = styles["align_right"]
        c10.number_format = "0.00%"

        # TPS
        c11 = ws.cell(row=cur_r, column=11, value=tx[10])
        c11.fill = fill_curr
        c11.font = styles["font_body"]
        c11.border = styles["cell_border"]
        c11.alignment = styles["align_right"]
        c11.number_format = "0.00"

        # Verdict
        c12 = ws.cell(row=cur_r, column=12, value=tx[11])
        c12.fill = styles["fill_pass"]
        c12.font = styles["font_pass"]
        c12.border = styles["cell_border"]
        c12.alignment = styles["align_center"]

        cur_r += 1

    end_tx_r = cur_r - 1

    # Total / Summary Row (Using Excel Formulas)
    ws.row_dimensions[cur_r].height = 24
    border_sum = Border(top=styles["border_thin"], bottom=styles["border_double"], left=styles["border_thin"], right=styles["border_thin"])

    c_tot_lbl = ws.cell(row=cur_r, column=1, value="総合計 / 加重平均 (Total)")
    c_tot_lbl.fill = styles["fill_sub"]
    c_tot_lbl.font = styles["font_body_bold"]
    c_tot_lbl.border = border_sum
    c_tot_lbl.alignment = styles["align_left"]

    # Sum of samples
    c_tot_sam = ws.cell(row=cur_r, column=2, value=f"=SUM(B{start_tx_r}:B{end_tx_r})")
    c_tot_sam.fill = styles["fill_sub"]
    c_tot_sam.font = styles["font_body_bold"]
    c_tot_sam.border = border_sum
    c_tot_sam.alignment = styles["align_right"]
    c_tot_sam.number_format = "#,##0"

    # Avg of avg
    c_tot_avg = ws.cell(row=cur_r, column=3, value=f"=AVERAGE(C{start_tx_r}:C{end_tx_r})")
    c_tot_avg.fill = styles["fill_sub"]
    c_tot_avg.font = styles["font_body_bold"]
    c_tot_avg.border = border_sum
    c_tot_avg.alignment = styles["align_right"]
    c_tot_avg.number_format = "#,##0"

    # Min of min
    c_tot_min = ws.cell(row=cur_r, column=4, value=f"=MIN(D{start_tx_r}:D{end_tx_r})")
    c_tot_min.fill = styles["fill_sub"]
    c_tot_min.font = styles["font_body_bold"]
    c_tot_min.border = border_sum
    c_tot_min.alignment = styles["align_right"]
    c_tot_min.number_format = "#,##0"

    # Max of max
    c_tot_max = ws.cell(row=cur_r, column=5, value=f"=MAX(E{start_tx_r}:E{end_tx_r})")
    c_tot_max.fill = styles["fill_sub"]
    c_tot_max.font = styles["font_body_bold"]
    c_tot_max.border = border_sum
    c_tot_max.alignment = styles["align_right"]
    c_tot_max.number_format = "#,##0"

    # Average of percentiles
    for c_i, col_let in enumerate(["F", "G", "H"], 6):
        c_p = ws.cell(row=cur_r, column=c_i, value=f"=AVERAGE({col_let}{start_tx_r}:{col_let}{end_tx_r})")
        c_p.fill = styles["fill_sub"]
        c_p.font = styles["font_body_bold"]
        c_p.border = border_sum
        c_p.alignment = styles["align_right"]
        c_p.number_format = "#,##0"

    # Total errors
    c_tot_err = ws.cell(row=cur_r, column=9, value=f"=SUM(I{start_tx_r}:I{end_tx_r})")
    c_tot_err.fill = styles["fill_sub"]
    c_tot_err.font = styles["font_body_bold"]
    c_tot_err.border = border_sum
    c_tot_err.alignment = styles["align_right"]
    c_tot_err.number_format = "#,##0"

    # Error % formula
    c_tot_ep = ws.cell(row=cur_r, column=10, value=f"=I{cur_r}/B{cur_r}")
    c_tot_ep.fill = styles["fill_sub"]
    c_tot_ep.font = styles["font_body_bold"]
    c_tot_ep.border = border_sum
    c_tot_ep.alignment = styles["align_right"]
    c_tot_ep.number_format = "0.00%"

    # Total TPS
    c_tot_tps = ws.cell(row=cur_r, column=11, value=f"=SUM(K{start_tx_r}:K{end_tx_r})")
    c_tot_tps.fill = styles["fill_sub"]
    c_tot_tps.font = styles["font_body_bold"]
    c_tot_tps.border = border_sum
    c_tot_tps.alignment = styles["align_right"]
    c_tot_tps.number_format = "0.00"

    # Total Verdict
    c_tot_v = ws.cell(row=cur_r, column=12, value="総合合格")
    c_tot_v.fill = styles["fill_pass"]
    c_tot_v.font = styles["font_pass"]
    c_tot_v.border = border_sum
    c_tot_v.alignment = styles["align_center"]

    cur_r += 2

    # Section 5: Server Resource Metrics
    styles["apply_section_header"](ws, cur_r, "5. サーバインフラリソース監視データ（CPU・メモリ・IOPS相関）", 12)
    cur_r += 1

    server_res = [
        ("負荷生成サーバ (JMeter CLI)", "4 vCPU / 8GB RAM", "18.5%", "34.2%", "2.4 GB", "30.0%", "120", "45 MB/s", "十分な余力あり。クライアント側ボトルネックなし。"),
        ("AWS API Gateway / ALB", "マネージド", "N/A", "N/A", "N/A", "N/A", "N/A", "180 MB/s", "レイテンシ増大なし、504タイムアウト極小。"),
        ("APサーバ (マイクロサービス群)", "8 vCPU / 16GB RAM", "42.8%", "68.5%", "8.2 GB", "51.2%", "850", "95 MB/s", "CPU使用率70%未満で健全に稼働。"),
        ("DBサーバ (Amazon Aurora MySQL)", "db.r6g.2xlarge (8vCPU/64GB)", "32.1%", "54.0%", "28.5 GB", "44.5%", "2,400", "30 MB/s", "コネクションプール(Max 100)枯渇なし、デッドロックなし。")
    ]

    tbl_headers7_res = ["対象サーバ・コンポーネント", "基本スペック", "平均CPU", "最大CPU", "使用メモリ", "メモリ使用率", "Disk IOPS", "ネットワーク", "インフラ評価コメント"]
    ws.row_dimensions[cur_r].height = 24
    cols_map7_res = [(1, 2), (3, 3), (4, 4), (5, 5), (6, 6), (7, 7), (8, 8), (9, 9), (10, 12)]
    for i, (sc, ec) in enumerate(cols_map7_res):
        if sc != ec:
            ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_tbl_header"]
            cell.border = styles["header_border"]
            if c == sc:
                cell.value = tbl_headers7_res[i]
                cell.font = styles["font_tbl_header"]
                cell.alignment = styles["align_center"]

    cur_r += 1
    for idx, r_item in enumerate(server_res):
        ws.row_dimensions[cur_r].height = 22
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # Name
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
        for c in range(1, 3):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = r_item[0]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Spec
        c_spec = ws.cell(row=cur_r, column=3, value=r_item[1])
        c_spec.fill = fill_curr
        c_spec.font = styles["font_body"]
        c_spec.border = styles["cell_border"]
        c_spec.alignment = styles["align_center"]

        # Avg CPU
        c_ac = ws.cell(row=cur_r, column=4, value=r_item[2])
        c_ac.fill = fill_curr
        c_ac.font = styles["font_body"]
        c_ac.border = styles["cell_border"]
        c_ac.alignment = styles["align_right"]

        # Max CPU
        c_mc = ws.cell(row=cur_r, column=5, value=r_item[3])
        c_mc.fill = fill_curr
        c_mc.font = styles["font_body_bold"]
        c_mc.border = styles["cell_border"]
        c_mc.alignment = styles["align_right"]

        # Mem Used
        c_mu = ws.cell(row=cur_r, column=6, value=r_item[4])
        c_mu.fill = fill_curr
        c_mu.font = styles["font_body"]
        c_mu.border = styles["cell_border"]
        c_mu.alignment = styles["align_right"]

        # Mem %
        c_mp = ws.cell(row=cur_r, column=7, value=r_item[5])
        c_mp.fill = fill_curr
        c_mp.font = styles["font_body"]
        c_mp.border = styles["cell_border"]
        c_mp.alignment = styles["align_right"]

        # IOPS
        c_io = ws.cell(row=cur_r, column=8, value=r_item[6])
        c_io.fill = fill_curr
        c_io.font = styles["font_body"]
        c_io.border = styles["cell_border"]
        c_io.alignment = styles["align_right"]

        # Net
        c_net = ws.cell(row=cur_r, column=9, value=r_item[7])
        c_net.fill = fill_curr
        c_net.font = styles["font_body"]
        c_net.border = styles["cell_border"]
        c_net.alignment = styles["align_right"]

        # Evaluation
        ws.merge_cells(start_row=cur_r, start_column=10, end_row=cur_r, end_column=12)
        for c in range(10, 13):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 10:
                c_cell.value = r_item[8]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]
        cur_r += 1

    # Section 6: Error Analysis
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "6. エラー分析・発生事象と根本原因・恒久対策 (Error Log Analysis)", 12)
    cur_r += 1

    errors = [
        ("2026/10/04 14:08:12", "06_注文決済確定", "HTTP 504 Gateway Timeout", 8, "決済サービス外部連携モックAPIの応答遅延 (タイムアウト3秒設定)", "外部決済APIクライアントのタイムアウト値を5秒に調整し、指数バックオフ付きリトライ機構を実装。"),
        ("2026/10/04 14:12:45", "03_商品一覧検索", "HTTP 500 Internal Error", 3, "同時アクセス集中時のDBコネクションプール一時的枯渇 (Max: 50到達)", "HikariCPコネクションプールの上限を50から100へ引き上げ、アイドルコネクション回収をチューニング。"),
        ("2026/10/04 14:15:30", "02_APIGW_CallResource", "HTTP 401 Unauthorized", 2, "Cognitoアクセストークンの有効期限失効 (1時間経過)", "トークン失効直前に自動リフレッシュするPreProcessorスクリプトをシナリオに追加。")
    ]

    tbl_headers7_err = ["発生時刻", "対象トランザクション", "HTTPコード / エラー種別", "件数", "根本原因 (Root Cause)", "恒久対策 / 是正措置"]
    ws.row_dimensions[cur_r].height = 24
    cols_map7_err = [(1, 2), (3, 4), (5, 6), (7, 7), (8, 9), (10, 12)]
    for i, (sc, ec) in enumerate(cols_map7_err):
        if sc != ec:
            ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_tbl_header"]
            cell.border = styles["header_border"]
            if c == sc:
                cell.value = tbl_headers7_err[i]
                cell.font = styles["font_tbl_header"]
                cell.alignment = styles["align_center"]

    cur_r += 1
    for idx, e in enumerate(errors):
        ws.row_dimensions[cur_r].height = 36
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # Time
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
        for c in range(1, 3):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = e[0]
                c_cell.font = styles["font_code"]
                c_cell.alignment = styles["align_center"]

        # Label
        ws.merge_cells(start_row=cur_r, start_column=3, end_row=cur_r, end_column=4)
        for c in range(3, 5):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 3:
                c_cell.value = e[1]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Code / Type
        ws.merge_cells(start_row=cur_r, start_column=5, end_row=cur_r, end_column=6)
        for c in range(5, 7):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 5:
                c_cell.value = e[2]
                c_cell.font = styles["font_fail"]
                c_cell.alignment = styles["align_left"]

        # Count
        c_cnt = ws.cell(row=cur_r, column=7, value=e[3])
        c_cnt.fill = fill_curr
        c_cnt.font = styles["font_body_bold"]
        c_cnt.border = styles["cell_border"]
        c_cnt.alignment = styles["align_right"]

        # Root Cause
        ws.merge_cells(start_row=cur_r, start_column=8, end_row=cur_r, end_column=9)
        for c in range(8, 10):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 8:
                c_cell.value = e[4]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        # Countermeasure
        ws.merge_cells(start_row=cur_r, start_column=10, end_row=cur_r, end_column=12)
        for c in range(10, 13):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 10:
                c_cell.value = e[5]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]
        cur_r += 1

    # Image Insertion Zones for Sheet 7 (Charts & Evidence)
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "7. レポート貼付用グラフ・ダッシュボード証跡エリア", 12)
    cur_r += 1
    styles["apply_image_placeholder"](
        ws, cur_r, 1, cur_r + 10, 6,
        "応答時間パーセンタイル推移グラフ (Response Time Percentiles)",
        "JMeter HTMLレポートまたはExcelで作成した、各トランザクションの90%/95%/99%パーセンタイル応答時間グラフ"
    )
    styles["apply_image_placeholder"](
        ws, cur_r, 7, cur_r + 10, 12,
        "スループット (TPS) & アクティブスレッド推移グラフ",
        "時間経過に伴うスループット (Throughput) と同時実行スレッド数の推移を重ね合わせたグラフ"
    )
    cur_r += 11

    styles["apply_image_placeholder"](
        ws, cur_r, 1, cur_r + 10, 12,
        "サーバインフラ監視ダッシュボード (Grafana / CloudWatch / Datadog)",
        "Web/AP/DBサーバのCPU使用率、メモリ推移、DBコネクション数、ネットワークIOの総合監視ダッシュボードキャプチャ"
    )
    cur_r += 11

    # Set column widths for Sheet 7
    col_widths7 = [16, 12, 11, 11, 11, 11, 11, 11, 11, 11, 12, 10]
    for idx, w in enumerate(col_widths7, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
