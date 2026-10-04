# -*- coding: utf-8 -*-
"""
Sheet 08: 確認観点と注意点 (Evaluation Points & Checklist)
"""

from openpyxl.utils import get_column_letter

def build_sheet_checklist(wb, styles):
    ws = wb.create_sheet(title="08_確認観点と注意点")
    ws.views.sheetView[0].showGridLines = True

    styles["apply_title_banner"](
        ws,
        "08. 負荷テスト現場の確認観点・陥りやすい落とし穴チェックリスト",
        "シナリオ設計・JMeter実行・Linux基盤・データ分析時のプロのチェック項目完全網羅",
        10
    )

    cur_r = 4
    styles["apply_section_header"](ws, cur_r, "1. 負荷テスト実施前・実施中・分析時の重要確認項目チェックリスト", 10)
    cur_r += 1

    checklist_items = [
        ("シナリオ設計", "思考時間 (Think Time) の設定", "リクエスト間に Uniform Random Timer 等で現実的な待機時間 (1〜3秒等) を設けているか？", "タイマーがない場合、実ユーザーの数万倍の負荷が局所集中し、非現実的なDDOS攻撃テストになってしまう。", "必須"),
        ("シナリオ設計", "Cookieとキャッシュの適切な設定", "HTTP Cookie Manager を配置し、Webブラウザと同等のセッション維持・キャッシュシミュレーションを行っているか？", "Cookieがないとセッションが維持できずログイン後の画面でエラー多発、または毎回新規セッション生成でAP負荷が異常増大する。", "必須"),
        ("シナリオ設計", "パラメータの多重化・CSV活用", "CSV Data Set Config を利用して、複数ユーザーのID/PASSや検索語を動的に変えているか？", "同一アカウント・同一IDで大量リクエストを送ると、DBの行ロック競合やアプリケーションの二重処理防止機能で不当にエラーになる。", "必須"),
        ("シナリオ設計", "段階的ランプアップの設定", "Ramp-up時間を十分に確保し、0秒から目標スレッド数まで徐々にアクセスを増加させているか？", "ランプアップ0秒で全スレッドを一斉起動すると、TCPコネクション確立やTLSハンドシェイクが瞬間飽和し、偽りのタイムアウトが発生する。", "強く推奨"),
        ("JMeterクライアント", "GUI実行の完全排除 (Non-GUI厳禁則)", "負荷テスト本番を必ずCLI (jmeter -n) で実行し、GUI画面から負荷をかけていないか？", "GUI実行はSwing描画スレッドがCPUを食いつぶし、JMeter自体がボトルネック化して正確なサーバ性能が測定できなくなる。", "絶対厳守"),
        ("JMeterクライアント", "高負荷リスナーの無効化・削除", "View Results Tree (結果をツリーで表示) 等の重いリスナーをJMXから削除または無効化しているか？", "View Results Treeが有効なままだと全リクエスト/レスポンスをメモリに溜め込み、JMeterが数分でOutOfMemoryError (OOM) 停止する。", "絶対厳守"),
        ("JMeterクライアント", "JVMヒープメモリの適正配分", "setenv.bat / setenv.sh で -Xms と -Xmx を同じ値にし、実メモリの50〜70%を割り当てているか？", "ヒープ不足は頻繁なFull GC (Stop The World) を引き起こし、JMeter側の停止によって見かけ上のレスポンス遅延やコネクション断が発生する。", "必須"),
        ("JMeterクライアント", "TCP Keep-Alive の設定整合性", "HTTP Request Defaults で「Use Keep-Alive」が有効になっているか？ (システム要件に準拠)", "Keep-Aliveが無効だとリクエストごとにTCP 3-wayハンドシェイクが発生し、TIME_WAITソケットの急増とポート枯渇を招く。", "必須"),
        ("Linux実行環境", "ファイルディスクリプタ上限 (ulimit)", "OSの ulimit -n が 65535 等の十分な値に拡張されているか？ (デフォルト1024のままになっていないか)", "スレッド数やコネクション数が増加した際、「java.io.IOException: Too many open files」エラーで全通信が即死する。", "必須"),
        ("Linux実行環境", "エフェメラルポート枯渇対策", "net.ipv4.ip_local_port_range の拡張および net.ipv4.tcp_tw_reuse = 1 が設定されているか？", "短時間に大量の接続・切断が発生するとポートが枯渇し、「Cannot assign requested address」エラーが発生する。", "必須"),
        ("測定・データ分析", "Coordinated Omission (協調的脱落) の理解", "サーバが一時的に遅延した際、JMeterのスレッドがブロックされてリクエスト送信数が減弾していないか？", "JMeterの標準スレッドは前リクエストの完了を待って次を送るため、サーバ遅延時に送信レートが落ち、見かけ上エラーが減る錯覚が起きる。", "重要観点"),
        ("測定・データ分析", "助走期間 (ウォームアップ) の除外", "テスト開始直後の初回接続・JITコンパイル・キャッシュ未暖気区間を分析データから切り離しているか？", "起動直後の初回トランザクションは極端に遅いため、定常状態のSLA評価に混入させるとデータが歪む。", "強く推奨"),
        ("測定・データ分析", "平均値の罠 (パーセンタイル重視の評価)", "平均値 (Average) だけで合否判定せず、90%ile / 95%ile / 99%ile の推移を分析しているか？", "平均値は少数の極端なロングテール遅延を隠蔽するため、「平均200msだが1%のユーザーは10秒待たされている」深刻な劣化を見落とす。", "必須"),
        ("測定・データ分析", "クライアント側リソースの健全性確認", "JMeter実行サーバ自体のCPU使用率が80%未満、メモリ使用率が70%未満に収まっていたか？", "負荷をかけているJMeterサーバ側がCPU100%に達していた場合、記録された遅延はサーバではなくJMeter自身の遅延である。", "絶対厳守")
    ]

    tbl_headers8 = ["カテゴリ", "確認項目 (観点)", "具体的な確認内容", "怠った場合に発生する重大トラブル・リスク", "重要度"]
    ws.row_dimensions[cur_r].height = 24
    cols_map8 = [(1, 2), (3, 4), (5, 7), (8, 9), (10, 10)]
    for i, (sc, ec) in enumerate(cols_map8):
        if sc != ec:
            ws.merge_cells(start_row=cur_r, start_column=sc, end_row=cur_r, end_column=ec)
        for c in range(sc, ec + 1):
            cell = ws.cell(row=cur_r, column=c)
            cell.fill = styles["fill_tbl_header"]
            cell.border = styles["header_border"]
            if c == sc:
                cell.value = tbl_headers8[i]
                cell.font = styles["font_tbl_header"]
                cell.alignment = styles["align_center"]

    cur_r += 1
    for idx, item in enumerate(checklist_items):
        ws.row_dimensions[cur_r].height = 42
        fill_curr = styles["fill_zebra"] if idx % 2 == 1 else styles["fill_white"]

        # Category
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=2)
        for c in range(1, 3):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 1:
                c_cell.value = item[0]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_center"]

        # Check Item
        ws.merge_cells(start_row=cur_r, start_column=3, end_row=cur_r, end_column=4)
        for c in range(3, 5):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 3:
                c_cell.value = item[1]
                c_cell.font = styles["font_body_bold"]
                c_cell.alignment = styles["align_left"]

        # Detail
        ws.merge_cells(start_row=cur_r, start_column=5, end_row=cur_r, end_column=7)
        for c in range(5, 8):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 5:
                c_cell.value = item[2]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        # Risk
        ws.merge_cells(start_row=cur_r, start_column=8, end_row=cur_r, end_column=9)
        for c in range(8, 10):
            c_cell = ws.cell(row=cur_r, column=c)
            c_cell.fill = fill_curr
            c_cell.border = styles["cell_border"]
            if c == 8:
                c_cell.value = item[3]
                c_cell.font = styles["font_body"]
                c_cell.alignment = styles["align_left"]

        # Priority
        p_cell = ws.cell(row=cur_r, column=10, value=item[4])
        p_cell.border = styles["cell_border"]
        p_cell.alignment = styles["align_center"]
        if item[4] in ["必須", "絶対厳守"]:
            p_cell.font = styles["font_fail"]
            p_cell.fill = styles["fill_fail"]
        else:
            p_cell.font = styles["font_warn"]
            p_cell.fill = styles["fill_warn"]
        cur_r += 1

    # Image Area for Checklist Evidence
    cur_r += 1
    styles["apply_section_header"](ws, cur_r, "2. 現場チェック完了証跡・レビュー承認印貼付エリア", 10)
    cur_r += 1
    styles["apply_image_placeholder"](
        ws, cur_r, 1, cur_r + 8, 10,
        "テスト実施前チェックリスト確認印 / レビュー承認証跡",
        "テスト設計者・インフラ担当者・プロジェクト責任者によるチェックリスト事前確認の承認証跡やサイン貼付欄"
    )
    cur_r += 9

    # Set column widths
    col_widths8 = [8, 14, 14, 16, 18, 18, 18, 18, 18, 12]
    for idx, w in enumerate(col_widths8, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
