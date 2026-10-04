# -*- coding: utf-8 -*-
"""
Common Styles and Helper Functions for JMeter Excel Workbook
Font: Meiryo UI everywhere
Palette: Professional Monotone Base
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def get_workbook_styles():
    FONT_NAME = "Meiryo UI"
    
    font_title = Font(name=FONT_NAME, size=15, bold=True, color="FFFFFF")
    font_section = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    font_sub = Font(name=FONT_NAME, size=10, bold=True, color="1F2421")
    font_tbl_header = Font(name=FONT_NAME, size=9.5, bold=True, color="FFFFFF")
    font_body = Font(name=FONT_NAME, size=9.5, color="212529")
    font_body_bold = Font(name=FONT_NAME, size=9.5, bold=True, color="212529")
    font_body_center = Font(name=FONT_NAME, size=9.5, color="212529")
    font_code = Font(name="Consolas", size=9, color="1A202C")
    font_note = Font(name=FONT_NAME, size=8.5, italic=True, color="4A5568")
    font_kpi_val = Font(name=FONT_NAME, size=14, bold=True, color="1A202C")
    font_kpi_lbl = Font(name=FONT_NAME, size=8.5, bold=True, color="4A5568")

    # Status Fonts
    font_pass = Font(name=FONT_NAME, size=9.5, bold=True, color="1B5E20")
    font_fail = Font(name=FONT_NAME, size=9.5, bold=True, color="B71C1C")
    font_warn = Font(name=FONT_NAME, size=9.5, bold=True, color="E65100")
    font_info = Font(name=FONT_NAME, size=9.5, bold=True, color="0D47A1")

    # Image Area Font
    font_img_title = Font(name=FONT_NAME, size=10, bold=True, color="374151")
    font_img_desc = Font(name=FONT_NAME, size=8.5, color="4B5563")

    # Fills (Monotone Base)
    fill_title = PatternFill(fill_type="solid", fgColor="2B303A")       # Dark Charcoal
    fill_section = PatternFill(fill_type="solid", fgColor="3E4452")     # Slate Gray
    fill_sub = PatternFill(fill_type="solid", fgColor="E8ECEF")         # Light Cool Gray
    fill_tbl_header = PatternFill(fill_type="solid", fgColor="4A5060")  # Table Header Slate
    fill_zebra = PatternFill(fill_type="solid", fgColor="F8F9FA")       # Zebra Light
    fill_white = PatternFill(fill_type="solid", fgColor="FFFFFF")
    fill_code = PatternFill(fill_type="solid", fgColor="EDF2F7")        # Code Block Gray
    fill_card = PatternFill(fill_type="solid", fgColor="F1F3F5")        # Card Box
    fill_img_area = PatternFill(fill_type="solid", fgColor="F3F4F6")    # Image paste area

    # Status Fills
    fill_pass = PatternFill(fill_type="solid", fgColor="D1E7DD")
    fill_fail = PatternFill(fill_type="solid", fgColor="F8D7DA")
    fill_warn = PatternFill(fill_type="solid", fgColor="FFF3CD")
    fill_info = PatternFill(fill_type="solid", fgColor="CFF4FC")

    # Borders
    border_thin = Side(style="thin", color="CBD5E0")
    border_medium = Side(style="medium", color="4A5060")
    border_double = Side(style="double", color="2B303A")
    border_dashed = Side(style="dashed", color="6B7280")

    cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    header_border = Border(left=border_thin, right=border_thin, top=border_medium, bottom=border_medium)
    sub_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    # Alignments
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    align_title = Alignment(horizontal="left", vertical="center", indent=1)
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

    styles = {
        "FONT_NAME": FONT_NAME,
        "font_title": font_title,
        "font_section": font_section,
        "font_sub": font_sub,
        "font_tbl_header": font_tbl_header,
        "font_body": font_body,
        "font_body_bold": font_body_bold,
        "font_body_center": font_body_center,
        "font_code": font_code,
        "font_note": font_note,
        "font_kpi_val": font_kpi_val,
        "font_kpi_lbl": font_kpi_lbl,
        "font_pass": font_pass,
        "font_fail": font_fail,
        "font_warn": font_warn,
        "font_info": font_info,
        "font_img_title": font_img_title,
        "font_img_desc": font_img_desc,
        "fill_title": fill_title,
        "fill_section": fill_section,
        "fill_sub": fill_sub,
        "fill_tbl_header": fill_tbl_header,
        "fill_zebra": fill_zebra,
        "fill_white": fill_white,
        "fill_code": fill_code,
        "fill_card": fill_card,
        "fill_img_area": fill_img_area,
        "fill_pass": fill_pass,
        "fill_fail": fill_fail,
        "fill_warn": fill_warn,
        "fill_info": fill_info,
        "border_thin": border_thin,
        "border_medium": border_medium,
        "border_double": border_double,
        "border_dashed": border_dashed,
        "cell_border": cell_border,
        "header_border": header_border,
        "sub_border": sub_border,
        "align_left": align_left,
        "align_center": align_center,
        "align_right": align_right,
        "align_title": align_title,
        "align_header": align_header,
    }

    def apply_title_banner(ws, title_text, subtitle_text, max_col=10):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        cell = ws.cell(row=1, column=1)
        cell.value = f" {title_text}"
        cell.font = font_title
        cell.fill = fill_title
        cell.alignment = align_title
        ws.row_dimensions[1].height = 36

        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
        cell_sub = ws.cell(row=2, column=1)
        cell_sub.value = f"  ■ {subtitle_text}"
        cell_sub.font = Font(name=FONT_NAME, size=9.5, italic=True, color="E2E8F0")
        cell_sub.fill = fill_section
        cell_sub.alignment = align_title
        ws.row_dimensions[2].height = 22

    def apply_section_header(ws, row_idx, section_title, max_col=10):
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=max_col)
        cell = ws.cell(row=row_idx, column=1)
        cell.value = f" ▶ {section_title}"
        cell.font = font_section
        cell.fill = fill_section
        cell.alignment = align_title
        ws.row_dimensions[row_idx].height = 26

    def apply_image_placeholder(ws, start_row, start_col, end_row, end_col, img_title, img_desc, recommended_size="幅800px × 高さ450px (16:9)"):
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.fill = fill_img_area
                top_s = border_dashed if r == start_row else None
                bot_s = border_dashed if r == end_row else None
                left_s = border_dashed if c == start_col else None
                right_s = border_dashed if c == end_col else None
                cell.border = Border(top=top_s, bottom=bot_s, left=left_s, right=right_s)

        ws.merge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)
        cell = ws.cell(row=start_row, column=start_col)
        text = (
            f"【画像・スクリーンショット貼付エリア】\n"
            f"● 対象: {img_title}\n"
            f"● 内容: {img_desc}\n"
            f"● 推奨サイズ: {recommended_size}\n"
            f"※ 貼付方法: Excel上部「挿入」リボン →「図」→「画像」→「このデバイス...」で画像を挿入し、この枠内にサイズ調整して配置してください。"
        )
        cell.value = text
        cell.font = font_img_desc
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    styles["apply_title_banner"] = apply_title_banner
    styles["apply_section_header"] = apply_section_header
    styles["apply_image_placeholder"] = apply_image_placeholder

    return styles
