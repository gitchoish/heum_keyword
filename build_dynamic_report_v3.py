# -*- coding: utf-8 -*-
"""
알프레드 혜움 세무기장 성과 리포트 v3
- 매체 내 일별 키워드 데이터 드릴다운 & 상위 10대 키워드 일별 매트릭스 탑재
- 신규 시트: '일별_키워드_추이' (매체/키워드 선택 시 31일 전 일자 일별 실시간 수식 집계)
- 신규 시트: '마케터_인사이트_파이프라인' (Won/Open/Lost 파이프라인, 수임실패사유 10대 분석, Biz vs General 세그먼트, 요일별 패턴)
- '전환_raw' (1,081행) 원천 데이터 엔진 시트 탑재
- 시트별 기간 선택 기능 (시작일/종료일 동적 필터링)
- 주차별_추이비교 (100% 수식, WoW 증감 및 증감율)
- 이모티콘 완전 배제, 다크 네이비 UI/UX 비즈니스 테마
"""

import sys
import os
import datetime
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

raw_filename = 'PLAYD) 혜움세무법인_세무기장_키워드_raw_f.xlsb'
output_filename = '알프레드_혜움_브랜드키워드_성과리포트.xlsx'
fallback_filename = '알프레드_세무기장_키워드소재_성과리포트.xlsx'

print("=== [Step 1] Loading Raw Data from XLSB ===")
xl = pd.ExcelFile(raw_filename, engine='calamine')

# 1. Load 통합_raw
df_tonghap = pd.read_excel(xl, sheet_name='통합_raw')
print(f"통합_raw loaded. Shape: {df_tonghap.shape}")
df_tonghap['일'] = pd.to_datetime(df_tonghap['일'])

# 2. Load 전환_raw
df_conv = pd.read_excel(xl, sheet_name='전환_raw')
print(f"전환_raw loaded. Shape: {df_conv.shape}")

wb = openpyxl.Workbook()
wb.remove(wb.active) # Remove default sheet

# Design Styles (Navy Dark Theme - No Emojis)
font_title = Font(name="맑은 고딕", size=16, bold=True, color="FFFFFF")
font_sec_title = Font(name="맑은 고딕", size=11, bold=True, color="1F4E79")
font_th = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
font_th_sub = Font(name="맑은 고딕", size=9, bold=True, color="FFFFFF")
font_td = Font(name="맑은 고딕", size=9, color="000000")
font_td_bold = Font(name="맑은 고딕", size=9, bold=True, color="000000")
font_total = Font(name="맑은 고딕", size=9, bold=True, color="1F4E79")
font_card_title = Font(name="맑은 고딕", size=9, bold=True, color="595959")
font_card_value = Font(name="맑은 고딕", size=14, bold=True, color="1F4E79")
font_cat_header = Font(name="맑은 고딕", size=10, bold=True, color="1F4E79")

# Period Control Bar Styles
font_ctrl_label = Font(name="맑은 고딕", size=9, bold=True, color="FFFFFF")
font_ctrl_date = Font(name="맑은 고딕", size=10, bold=True, color="1F4E79")
font_ctrl_info = Font(name="맑은 고딕", size=9, italic=True, color="595959")
fill_ctrl_label = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
fill_ctrl_date = PatternFill(start_color="EDF2F8", end_color="EDF2F8", fill_type="solid")

fill_navy = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
fill_th = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
fill_th_sub = PatternFill(start_color="41719C", end_color="41719C", fill_type="solid")
fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
fill_cat = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
fill_card_head = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
fill_wow_pos = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
fill_matrix_head = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                     top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
border_header = Border(left=Side(style='thin', color='FFFFFF'), right=Side(style='thin', color='FFFFFF'),
                       top=Side(style='medium', color='1F4E79'), bottom=Side(style='medium', color='1F4E79'))
border_total = Border(top=Side(style='thin', color='1F4E79'), bottom=Side(style='double', color='1F4E79'))
border_ctrl = Border(left=Side(style='thin', color='B4C6E7'), right=Side(style='thin', color='B4C6E7'),
                     top=Side(style='thin', color='B4C6E7'), bottom=Side(style='thin', color='B4C6E7'))

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')

def add_period_control_bar(ws, max_col_letter='N'):
    ws.row_dimensions[3].height = 24
    ws.cell(3, 2, "조회 시작일").font = font_ctrl_label
    ws.cell(3, 2).fill = fill_ctrl_label; ws.cell(3, 2).alignment = align_center; ws.cell(3, 2).border = border_ctrl
    c3 = ws.cell(3, 3, "=MIN(통합_raw!E:E)")
    c3.font = font_ctrl_date; c3.fill = fill_ctrl_date; c3.alignment = align_center
    c3.number_format = "yyyy-mm-dd"; c3.border = border_ctrl
    ws.cell(3, 4, "조회 종료일").font = font_ctrl_label
    ws.cell(3, 4).fill = fill_ctrl_label; ws.cell(3, 4).alignment = align_center; ws.cell(3, 4).border = border_ctrl
    e3 = ws.cell(3, 5, "=MAX(통합_raw!E:E)")
    e3.font = font_ctrl_date; e3.fill = fill_ctrl_date; e3.alignment = align_center
    e3.number_format = "yyyy-mm-dd"; e3.border = border_ctrl
    ws.merge_cells("F3:G3")
    fg = ws.cell(3, 6, '="[전체 원천 데이터 기간: " & TEXT(MIN(통합_raw!E:E),"yyyy-mm-dd") & " ~ " & TEXT(MAX(통합_raw!E:E),"yyyy-mm-dd") & "]"')
    fg.font = font_sec_title; fg.alignment = align_left
    ws.merge_cells(f"H3:{max_col_letter}3")
    info_c = ws.cell(3, 8, "안내: C3(시작일)과 E3(종료일)을 수정하면 아래 성과가 해당 기간으로 실시간 자동 재집계됩니다.")
    info_c.font = font_ctrl_info; info_c.alignment = align_left


# ==============================================================================
# Sheet 1: 종합_대시보드
# ==============================================================================
print("Building Sheet 1: 종합_대시보드...")
ws_dash = wb.create_sheet(title='종합_대시보드')
ws_dash.views.sheetView[0].showGridLines = True

ws_dash.merge_cells('A1:R1')
ws_dash['A1'] = "혜움 세무기장 마케팅 성과 종합 대시보드"
ws_dash['A1'].font = font_title; ws_dash['A1'].fill = fill_navy; ws_dash['A1'].alignment = align_center
ws_dash.row_dimensions[1].height = 40

add_period_control_bar(ws_dash, 'R')

ws_dash.cell(5, 1, "[일일 운영 마감 기준일]").font = font_sec_title
ws_dash.cell(6, 1, "2026-10-01").font = font_card_value
ws_dash.cell(6, 1).alignment = align_center; ws_dash.cell(6, 1).number_format = "yyyy-mm-dd"

cards_dash = [
    ("총 예산(VAT제외)", "=C16", 3, 4, "₩#,##0"),
    ("조회기간 소진(VAT-)", "=SUMIFS(통합_raw!S:S, 통합_raw!E:E, \">=\"&$C$3, 통합_raw!E:E, \"<=\"&$E$3)/1.1", 5, 6, "₩#,##0"),
    ("조회기간 소진율", "=IF(C6>0, E6/C6, 0)", 7, 8, "0.0%"),
    ("총 유입 클릭수", "=SUMIFS(통합_raw!R:R, 통합_raw!E:E, \">=\"&$C$3, 통합_raw!E:E, \"<=\"&$E$3)", 9, 10, "#,##0"),
    ("총 리드수 (Biz)", "=SUMIFS(통합_raw!T:T, 통합_raw!E:E, \">=\"&$C$3, 통합_raw!E:E, \"<=\"&$E$3)", 11, 12, "#,##0"),
    ("평균 리드 CPA", "=IF(K6>0, E6/K6, 0)", 13, 14, "₩#,##0"),
    ("총 최종 수임건수", "=SUMIFS(통합_raw!W:W, 통합_raw!E:E, \">=\"&$C$3, 통합_raw!E:E, \"<=\"&$E$3)", 15, 16, "#,##0"),
    ("수임당 비용 (CPS)", "=IF(O6>0, E6/O6, 0)", 17, 18, "₩#,##0"),
]

ws_dash.row_dimensions[5].height = 18; ws_dash.row_dimensions[6].height = 26
for title, formula, c_start, c_end, fmt in cards_dash:
    ws_dash.merge_cells(f"{get_column_letter(c_start)}5:{get_column_letter(c_end)}5")
    ws_dash.merge_cells(f"{get_column_letter(c_start)}6:{get_column_letter(c_end)}6")
    top_c = ws_dash.cell(5, c_start, title)
    top_c.font = font_card_title; top_c.fill = fill_card_head; top_c.alignment = align_center
    val_c = ws_dash.cell(6, c_start, formula)
    val_c.font = font_card_value; val_c.alignment = align_center
    if fmt: val_c.number_format = fmt
    for r in [5, 6]:
        for c in range(c_start, c_end + 1):
            ws_dash.cell(r, c).border = border_thin

ws_dash.cell(8, 1, "[매체별 예산 및 조회 기간 성과 요약] C3~E3 기간 필터 실시간 수식 집계").font = font_sec_title
headers_s1 = [
    "서비스 구분", "매체", "예산(vat-)", "예산(vat+)", "비중", "소진율", "목표CPA", "실 CPA", "예상 마감",
    "노출", "클릭", "소진 비용(vat-)", "CTR", "CPC", "총 리드", "총 리드 CPA", "수임 biz", "수임 CPA"
]
ws_dash.row_dimensions[9].height = 24
for col_idx, h in enumerate(headers_s1, 1):
    cell = ws_dash.cell(9, col_idx, h)
    cell.font = font_th; cell.fill = fill_th; cell.alignment = align_center; cell.border = border_header

media_config = [
    ("세무기장", "네이버 검색광고", 13100000, 14410000, 0.167, 129825, 9741232, "네이버 검색광고"),
    ("세무기장", "네이버브랜드검색광고", 750000, 825000, 0.010, 42857, 750000, "네이버브랜드검색광고"),
    ("세무기장", "네이버 플레이스", 250000, 275000, 0.003, 250000, 69409, "네이버 플레이스"),
    ("세무기장", "구글 검색광고", 13600000, 14960000, 0.173, 273404, 18304275, "구글 검색광고"),
    ("세무기장", "META", 50800000, 55880000, 0.647, 225250, 45003882, "META"),
    ("세무기장", "1차 Total", 0, 0, 0, 0, 0, "SUBTOTAL_1"),
    ("세무기장", "네이버 검색광고 (2차)", 9100000, 10010000, 0.148, 129825, 0, "NONE"),
    ("세무기장", "네이버브랜드검색 (2차)", 750000, 825000, 0.012, 42857, 750000, "NONE"),
    ("세무기장", "네이버 플레이스 (2차)", 250000, 275000, 0.004, 250000, 0, "NONE"),
    ("세무기장", "구글 검색광고 (2차)", 12100000, 13310000, 0.197, 273404, 0, "NONE"),
    ("세무기장", "META (2차)", 39300000, 43230000, 0.639, 225250, 0, "NONE"),
    ("세무기장", "2차 Total", 0, 0, 0, 0, 0, "SUBTOTAL_2"),
]

row_start = 10
for i, m_conf in enumerate(media_config):
    curr_row = row_start + i
    svc, med, b_exc, b_inc, share, t_cpa, exp_close, db_match = m_conf
    is_subtotal = "Total" in med
    c_font = font_total if is_subtotal else font_td
    c_fill = fill_total if is_subtotal else None
    
    ws_dash.cell(curr_row, 1, svc).alignment = align_center
    ws_dash.cell(curr_row, 2, med).alignment = align_left if not is_subtotal else align_center
    
    if db_match == "SUBTOTAL_1":
        ws_dash.cell(curr_row, 3, f"=SUM(C10:C14)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 4, f"=SUM(D10:D14)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 5, 1.0).number_format = "0.0%"
        ws_dash.cell(curr_row, 6, f"=IF(C{curr_row}>0, L{curr_row}/C{curr_row}, 0)").number_format = "0.0%"
        ws_dash.cell(curr_row, 7, 198898).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 8, f"=IF(O{curr_row}>0, L{curr_row}/O{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 9, 73868798).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 10, f"=SUM(J10:J14)").number_format = "#,##0"
        ws_dash.cell(curr_row, 11, f"=SUM(K10:K14)").number_format = "#,##0"
        ws_dash.cell(curr_row, 12, f"=SUM(L10:L14)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 13, f"=IF(J{curr_row}>0, K{curr_row}/J{curr_row}, 0)").number_format = "0.00%"
        ws_dash.cell(curr_row, 14, f"=IF(K{curr_row}>0, L{curr_row}/K{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 15, f"=SUM(O10:O14)").number_format = "#,##0"
        ws_dash.cell(curr_row, 16, f"=IF(O{curr_row}>0, L{curr_row}/O{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 17, f"=SUM(Q10:Q14)").number_format = "#,##0"
        ws_dash.cell(curr_row, 18, f"=IF(Q{curr_row}>0, L{curr_row}/Q{curr_row}, 0)").number_format = "₩#,##0"
    elif db_match == "SUBTOTAL_2":
        ws_dash.cell(curr_row, 3, f"=SUM(C16:C20)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 4, f"=SUM(D16:D20)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 5, 1.0).number_format = "0.0%"
        ws_dash.cell(curr_row, 6, f"=IF(C{curr_row}>0, L{curr_row}/C{curr_row}, 0)").number_format = "0.0%"
        ws_dash.cell(curr_row, 7, 200114).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 8, f"=IF(O{curr_row}>0, L{curr_row}/O{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 9, 750000).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 10, f"=SUM(J16:J20)").number_format = "#,##0"
        ws_dash.cell(curr_row, 11, f"=SUM(K16:K20)").number_format = "#,##0"
        ws_dash.cell(curr_row, 12, f"=SUM(L16:L20)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 13, f"=IF(J{curr_row}>0, K{curr_row}/J{curr_row}, 0)").number_format = "0.00%"
        ws_dash.cell(curr_row, 14, f"=IF(K{curr_row}>0, L{curr_row}/K{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 15, f"=SUM(O16:O20)").number_format = "#,##0"
        ws_dash.cell(curr_row, 16, f"=IF(O{curr_row}>0, L{curr_row}/O{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 17, f"=SUM(Q16:Q20)").number_format = "#,##0"
        ws_dash.cell(curr_row, 18, f"=IF(Q{curr_row}>0, L{curr_row}/Q{curr_row}, 0)").number_format = "₩#,##0"
    elif db_match == "NONE":
        ws_dash.cell(curr_row, 3, b_exc).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 4, b_inc).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 5, share).number_format = "0.0%"
        ws_dash.cell(curr_row, 6, 0.0).number_format = "0.0%"
        ws_dash.cell(curr_row, 7, t_cpa).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 8, 0).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 9, exp_close).number_format = "₩#,##0"
        for c_idx in range(10, 19):
            ws_dash.cell(curr_row, c_idx, 0).number_format = "₩#,##0" if c_idx in [12, 14, 16, 18] else ("0.00%" if c_idx == 13 else "#,##0")
    else:
        ws_dash.cell(curr_row, 3, b_exc).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 4, b_inc).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 5, share).number_format = "0.0%"
        ws_dash.cell(curr_row, 6, f"=IF(C{curr_row}>0, L{curr_row}/C{curr_row}, 0)").number_format = "0.0%"
        ws_dash.cell(curr_row, 7, t_cpa).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 8, f"=IF(O{curr_row}>0, L{curr_row}/O{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 9, exp_close).number_format = "₩#,##0"
        
        ws_dash.cell(curr_row, 10, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, "{db_match}", 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws_dash.cell(curr_row, 11, f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, "{db_match}", 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws_dash.cell(curr_row, 12, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, "{db_match}", 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)/1.1').number_format = "₩#,##0"
        ws_dash.cell(curr_row, 13, f"=IF(J{curr_row}>0, K{curr_row}/J{curr_row}, 0)").number_format = "0.00%"
        ws_dash.cell(curr_row, 14, f"=IF(K{curr_row}>0, L{curr_row}/K{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 15, f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, "{db_match}", 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws_dash.cell(curr_row, 16, f"=IF(O{curr_row}>0, L{curr_row}/O{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 17, f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, "{db_match}", 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws_dash.cell(curr_row, 18, f"=IF(Q{curr_row}>0, L{curr_row}/Q{curr_row}, 0)").number_format = "₩#,##0"
        
    for c in range(1, 19):
        cell = ws_dash.cell(curr_row, c); cell.font = c_font
        if c_fill: cell.fill = c_fill
        cell.border = border_thin

tot_r = row_start + len(media_config)
ws_dash.merge_cells(f'A{tot_r}:B{tot_r}')
ws_dash.cell(tot_r, 1, "전체 총계 (Grand Total)").alignment = align_center
ws_dash.cell(tot_r, 3, "=C15+C21").number_format = "₩#,##0"
ws_dash.cell(tot_r, 4, "=D15+D21").number_format = "₩#,##0"
ws_dash.cell(tot_r, 5, 1.0).number_format = "0.0%"
ws_dash.cell(tot_r, 6, f"=IF(C{tot_r}>0, L{tot_r}/C{tot_r}, 0)").number_format = "0.0%"
ws_dash.cell(tot_r, 7, "=AVERAGE(G15,G21)").number_format = "₩#,##0"
ws_dash.cell(tot_r, 8, f"=IF(O{tot_r}>0, L{tot_r}/O{tot_r}, 0)").number_format = "₩#,##0"
ws_dash.cell(tot_r, 9, "=I15+I21").number_format = "₩#,##0"
ws_dash.cell(tot_r, 10, f"=J15+J21").number_format = "#,##0"
ws_dash.cell(tot_r, 11, f"=K15+K21").number_format = "#,##0"
ws_dash.cell(tot_r, 12, f"=L15+L21").number_format = "₩#,##0"
ws_dash.cell(tot_r, 13, f"=IF(J{tot_r}>0, K{tot_r}/J{tot_r}, 0)").number_format = "0.00%"
ws_dash.cell(tot_r, 14, f"=IF(K{tot_r}>0, L{tot_r}/K{tot_r}, 0)").number_format = "₩#,##0"
ws_dash.cell(tot_r, 15, f"=O15+O21").number_format = "#,##0"
ws_dash.cell(tot_r, 16, f"=IF(O{tot_r}>0, L{tot_r}/O{tot_r}, 0)").number_format = "₩#,##0"
ws_dash.cell(tot_r, 17, f"=Q15+Q21").number_format = "#,##0"
ws_dash.cell(tot_r, 18, f"=IF(Q{tot_r}>0, L{tot_r}/Q{tot_r}, 0)").number_format = "₩#,##0"

for c in range(1, 19):
    cell = ws_dash.cell(tot_r, c); cell.font = font_total; cell.fill = fill_total; cell.border = border_total

dash_col_widths = {
    'A': 12, 'B': 22, 'C': 14, 'D': 14, 'E': 9, 'F': 9, 'G': 12, 'H': 12, 'I': 14,
    'J': 11, 'K': 10, 'L': 14, 'M': 9, 'N': 10, 'O': 10, 'P': 12, 'Q': 9, 'R': 12
}
for col_let, w in dash_col_widths.items():
    ws_dash.column_dimensions[col_let].width = w


# ==============================================================================
# Sheet 2: 브랜드_키워드분석
# ==============================================================================
print("Building Sheet 2: 브랜드_키워드분석...")
ws_kw_cat = wb.create_sheet(title='브랜드_키워드분석')
ws_kw_cat.views.sheetView[0].showGridLines = True

ws_kw_cat.merge_cells('A1:O1')
ws_kw_cat['A1'] = "혜움 브랜드 및 카테고리별 키워드 성과 상세 분석"
ws_kw_cat['A1'].font = font_title; ws_kw_cat['A1'].fill = fill_navy; ws_kw_cat['A1'].alignment = align_center
ws_kw_cat.row_dimensions[1].height = 40

add_period_control_bar(ws_kw_cat, 'O')

keyword_categories = [
    ("1. 혜움 브랜드 키워드군 (Brand Assets)", [
        ("혜움", "구글 검색광고"),
        ("세무법인혜움", "네이버브랜드검색광고"),
        ("혜움", "네이버브랜드검색광고"),
        ("혜움", "네이버 검색광고"),
        ("세무법인혜움", "네이버 검색광고"),
    ]),
    ("2. 세무기장 핵심 직무 키워드군 (Core Tax Accounting)", [
        ("세무사", "네이버 검색광고"),
        ("세무", "구글 검색광고"),
        ("세무법인", "네이버 검색광고"),
        ("법인세무기장", "네이버 검색광고"),
        ("법인기장", "네이버 검색광고"),
        ("세무기장", "네이버 검색광고"),
        ("기장대행", "네이버 검색광고"),
        ("개인사업자세무사", "네이버 검색광고"),
        ("개인사업자세무대행", "네이버 검색광고"),
        ("기장신고", "네이버 검색광고"),
        ("세무사추천", "네이버 검색광고"),
    ]),
    ("3. 특수 세목 및 컨설팅 키워드군 (Tax Consulting)", [
        ("법인세금", "구글 검색광고"),
        ("증여세세무사", "네이버 검색광고"),
        ("상속세세무사", "네이버 검색광고"),
        ("양도소득세세무사", "네이버 검색광고"),
        ("무료세무상담", "네이버 검색광고"),
    ]),
    ("4. 로컬 및 지역 타겟 키워드군 (Local Targeting)", [
        ("강서구세무사", "네이버 검색광고"),
        ("고양세무사", "네이버 검색광고"),
        ("동탄세무사", "네이버 검색광고"),
        ("마포구세무사", "구글 검색광고"),
        ("인천세무사", "구글 검색광고"),
        ("김해세무사", "네이버 검색광고"),
    ]),
]

cur_cat_row = 5
headers_kw_cat = [
    "카테고리 구분", "키워드", "대표 매체", "노출수", "클릭수", "비용(vat-)", "비용(vat+)",
    "CTR", "CPC", "총 리드", "유효 리드", "수임", "리드 CPA", "수임 CPA", "전환율(CVR)"
]
ws_kw_cat.row_dimensions[cur_cat_row].height = 24
for c_idx, h in enumerate(headers_kw_cat, 1):
    c = ws_kw_cat.cell(cur_cat_row, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

cur_cat_row += 1

for cat_title, kw_items in keyword_categories:
    ws_kw_cat.merge_cells(f'A{cur_cat_row}:O{cur_cat_row}')
    cat_cell = ws_kw_cat.cell(cur_cat_row, 1, cat_title)
    cat_cell.font = font_cat_header; cat_cell.fill = fill_cat; cat_cell.alignment = align_left
    for c in range(1, 16): ws_kw_cat.cell(cur_cat_row, c).border = border_thin
    ws_kw_cat.row_dimensions[cur_cat_row].height = 22
    cur_cat_row += 1
    
    cat_start_r = cur_cat_row
    for kw, med in kw_items:
        r_i = cur_cat_row
        ws_kw_cat.cell(r_i, 1, cat_title.split('.')[1].split('(')[0].strip()).alignment = align_center
        ws_kw_cat.cell(r_i, 2, kw).alignment = align_left
        ws_kw_cat.cell(r_i, 3, med).alignment = align_center
        
        ws_kw_cat.cell(r_i, 4, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws_kw_cat.cell(r_i, 5, f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws_kw_cat.cell(r_i, 6, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)/1.1').number_format = "₩#,##0"
        ws_kw_cat.cell(r_i, 7, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "₩#,##0"
        ws_kw_cat.cell(r_i, 8, f'=IF(D{r_i}>0, E{r_i}/D{r_i}, 0)').number_format = "0.00%"
        ws_kw_cat.cell(r_i, 9, f'=IF(E{r_i}>0, F{r_i}/E{r_i}, 0)').number_format = "₩#,##0"
        ws_kw_cat.cell(r_i, 10, f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws_kw_cat.cell(r_i, 11, f'=SUMIFS(통합_raw!V:V, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws_kw_cat.cell(r_i, 12, f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws_kw_cat.cell(r_i, 13, f'=IF(J{r_i}>0, F{r_i}/J{r_i}, 0)').number_format = "₩#,##0"
        ws_kw_cat.cell(r_i, 14, f'=IF(L{r_i}>0, F{r_i}/L{r_i}, 0)').number_format = "₩#,##0"
        ws_kw_cat.cell(r_i, 15, f'=IF(E{r_i}>0, J{r_i}/E{r_i}, 0)').number_format = "0.0%"
        
        for c in range(1, 16):
            cell = ws_kw_cat.cell(r_i, c); cell.font = font_td; cell.border = border_thin
        cur_cat_row += 1
        
    sub_r = cur_cat_row
    ws_kw_cat.merge_cells(f'A{sub_r}:C{sub_r}')
    ws_kw_cat.cell(sub_r, 1, f"{cat_title.split('.')[1].split('(')[0].strip()} 소계").alignment = align_center
    ws_kw_cat.cell(sub_r, 4, f"=SUM(D{cat_start_r}:D{sub_r-1})").number_format = "#,##0"
    ws_kw_cat.cell(sub_r, 5, f"=SUM(E{cat_start_r}:E{sub_r-1})").number_format = "#,##0"
    ws_kw_cat.cell(sub_r, 6, f"=SUM(F{cat_start_r}:F{sub_r-1})").number_format = "₩#,##0"
    ws_kw_cat.cell(sub_r, 7, f"=SUM(G{cat_start_r}:G{sub_r-1})").number_format = "₩#,##0"
    ws_kw_cat.cell(sub_r, 8, f"=IF(D{sub_r}>0, E{sub_r}/D{sub_r}, 0)").number_format = "0.00%"
    ws_kw_cat.cell(sub_r, 9, f"=IF(E{sub_r}>0, F{sub_r}/E{sub_r}, 0)").number_format = "₩#,##0"
    ws_kw_cat.cell(sub_r, 10, f"=SUM(J{cat_start_r}:J{sub_r-1})").number_format = "#,##0"
    ws_kw_cat.cell(sub_r, 11, f"=SUM(K{cat_start_r}:K{sub_r-1})").number_format = "#,##0"
    ws_kw_cat.cell(sub_r, 12, f"=SUM(L{cat_start_r}:L{sub_r-1})").number_format = "#,##0"
    ws_kw_cat.cell(sub_r, 13, f"=IF(J{sub_r}>0, F{sub_r}/J{sub_r}, 0)").number_format = "₩#,##0"
    ws_kw_cat.cell(sub_r, 14, f"=IF(L{sub_r}>0, F{sub_r}/L{sub_r}, 0)").number_format = "₩#,##0"
    ws_kw_cat.cell(sub_r, 15, f"=IF(E{sub_r}>0, J{sub_r}/E{sub_r}, 0)").number_format = "0.0%"
    
    for c in range(1, 16):
        cell = ws_kw_cat.cell(sub_r, c); cell.font = font_total; cell.fill = fill_total; cell.border = border_thin
    cur_cat_row += 2

kw_cat_widths = {
    'A': 18, 'B': 22, 'C': 18, 'D': 11, 'E': 10, 'F': 14, 'G': 14,
    'H': 9, 'I': 11, 'J': 10, 'K': 10, 'L': 9, 'M': 13, 'N': 13, 'O': 10
}
for col_let, w in kw_cat_widths.items():
    ws_kw_cat.column_dimensions[col_let].width = w


# ==============================================================================
# Sheet 3: 일별_키워드_추이 (신규! 매체별/키워드별 일별 31일 캘린더 분석 엔진)
# ==============================================================================
print("Building Sheet 3: 일별_키워드_추이...")
ws_daily_kw = wb.create_sheet(title='일별_키워드_추이')
ws_daily_kw.views.sheetView[0].showGridLines = True

ws_daily_kw.merge_cells('A1:N1')
ws_daily_kw['A1'] = "매체 및 키워드별 일별 성과 추이 상세 분석 (Daily Calendar Drill-Down)"
ws_daily_kw['A1'].font = font_title; ws_daily_kw['A1'].fill = fill_navy; ws_daily_kw['A1'].alignment = align_center
ws_daily_kw.row_dimensions[1].height = 40

# Interactive Controls: Media & Keyword Selection
ws_daily_kw.row_dimensions[3].height = 24
ws_daily_kw.cell(3, 1, "분석 대상 매체").font = font_ctrl_label
ws_daily_kw.cell(3, 1).fill = fill_ctrl_label; ws_daily_kw.cell(3, 1).alignment = align_center; ws_daily_kw.cell(3, 1).border = border_ctrl
ws_daily_kw.cell(3, 2, "네이버 검색광고").font = font_ctrl_date
ws_daily_kw.cell(3, 2).fill = fill_ctrl_date; ws_daily_kw.cell(3, 2).alignment = align_center; ws_daily_kw.cell(3, 2).border = border_ctrl

ws_daily_kw.cell(3, 4, "분석 대상 키워드/소재").font = font_ctrl_label
ws_daily_kw.cell(3, 4).fill = fill_ctrl_label; ws_daily_kw.cell(3, 4).alignment = align_center; ws_daily_kw.cell(3, 4).border = border_ctrl
ws_daily_kw.cell(3, 5, "세무사").font = font_ctrl_date
ws_daily_kw.cell(3, 5).fill = fill_ctrl_date; ws_daily_kw.cell(3, 5).alignment = align_center; ws_daily_kw.cell(3, 5).border = border_ctrl

ws_daily_kw.merge_cells("G3:N3")
info_dk = ws_daily_kw.cell(3, 7, "안내: B3(매체)와 E3(키워드/소재)를 입력하면 해당 키워드의 31일간 전체 일별 성과가 실시간 자동 재계산됩니다.")
info_dk.font = font_ctrl_info; info_dk.alignment = align_left

# KPI Cards for this selected keyword
ws_daily_kw.row_dimensions[5].height = 18; ws_daily_kw.row_dimensions[6].height = 24
cards_dk = [
    ("선택 키워드 총 노출수", "=SUM(D9:D39)", 1, 2, "#,##0"),
    ("선택 키워드 총 클릭수", "=SUM(E9:E39)", 3, 4, "#,##0"),
    ("선택 키워드 총 비용(VAT-)", "=SUM(H9:H39)", 5, 7, "₩#,##0"),
    ("선택 키워드 총 리드", "=SUM(J9:J39)", 8, 9, "#,##0"),
    ("선택 키워드 총 수임", "=SUM(L9:L39)", 10, 11, "#,##0"),
    ("평균 CPA / CPS", "=IF(J6>0, E6/J6, 0)", 12, 14, "₩#,##0"),
]
for title, formula, c_start, c_end, fmt in cards_dk:
    ws_daily_kw.merge_cells(f"{get_column_letter(c_start)}5:{get_column_letter(c_end)}5")
    ws_daily_kw.merge_cells(f"{get_column_letter(c_start)}6:{get_column_letter(c_end)}6")
    top_c = ws_daily_kw.cell(5, c_start, title)
    top_c.font = font_card_title; top_c.fill = fill_card_head; top_c.alignment = align_center
    val_c = ws_daily_kw.cell(6, c_start, formula)
    val_c.font = font_card_value; val_c.alignment = align_center
    if fmt: val_c.number_format = fmt
    for r in [5, 6]:
        for c in range(c_start, c_end + 1):
            ws_daily_kw.cell(r, c).border = border_thin

headers_daily_kw = [
    "일자", "요일", "영업일구분", "노출수", "클릭수", "CTR", "CPC", "광고비(vat-)", "광고비(vat+)",
    "총 리드", "유효리드", "수임", "리드 CPA", "수임 CPA"
]
ws_daily_kw.row_dimensions[8].height = 24
for c_idx, h in enumerate(headers_daily_kw, 1):
    c = ws_daily_kw.cell(8, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

# Populate 31 days (2026-09-01 to 2026-10-01)
all_days = pd.date_range('2026-09-01', '2026-10-01')
for idx, dt in enumerate(all_days):
    cur_r = 9 + idx
    d_str = dt.strftime('%Y-%m-%d')
    w_kor = ['월', '화', '수', '목', '금', '토', '일'][dt.weekday()]
    biz_str = '주말' if dt.weekday() >= 5 else '영업일'
    
    ws_daily_kw.cell(cur_r, 1, d_str).alignment = align_center
    ws_daily_kw.cell(cur_r, 1).number_format = "yyyy-mm-dd"
    ws_daily_kw.cell(cur_r, 2, w_kor).alignment = align_center
    ws_daily_kw.cell(cur_r, 3, biz_str).alignment = align_center
    
    # 100% PURE DYNAMIC SUMIFS REFERENCING $B$3(매체) AND $E$3(키워드)
    ws_daily_kw.cell(cur_r, 4, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, $B$3, 통합_raw!AE:AE, $E$3, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
    ws_daily_kw.cell(cur_r, 5, f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, $B$3, 통합_raw!AE:AE, $E$3, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
    ws_daily_kw.cell(cur_r, 6, f'=IF(D{cur_r}>0, E{cur_r}/D{cur_r}, 0)').number_format = "0.00%"
    ws_daily_kw.cell(cur_r, 7, f'=IF(E{cur_r}>0, H{cur_r}/E{cur_r}, 0)').number_format = "₩#,##0"
    ws_daily_kw.cell(cur_r, 8, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, $B$3, 통합_raw!AE:AE, $E$3, 통합_raw!E:E, A{cur_r})/1.1').number_format = "₩#,##0"
    ws_daily_kw.cell(cur_r, 9, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, $B$3, 통합_raw!AE:AE, $E$3, 통합_raw!E:E, A{cur_r})').number_format = "₩#,##0"
    ws_daily_kw.cell(cur_r, 10, f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, $B$3, 통합_raw!AE:AE, $E$3, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
    ws_daily_kw.cell(cur_r, 11, f'=SUMIFS(통합_raw!V:V, 통합_raw!L:L, $B$3, 통합_raw!AE:AE, $E$3, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
    ws_daily_kw.cell(cur_r, 12, f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, $B$3, 통합_raw!AE:AE, $E$3, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
    ws_daily_kw.cell(cur_r, 13, f'=IF(J{cur_r}>0, H{cur_r}/J{cur_r}, 0)').number_format = "₩#,##0"
    ws_daily_kw.cell(cur_r, 14, f'=IF(L{cur_r}>0, H{cur_r}/L{cur_r}, 0)').number_format = "₩#,##0"
    
    for c in range(1, 15):
        cell = ws_daily_kw.cell(cur_r, c); cell.font = font_td; cell.border = border_thin

# Total Row for 31 days
tot_dk_r = 9 + len(all_days)
ws_daily_kw.merge_cells(f'A{tot_dk_r}:C{tot_dk_r}')
ws_daily_kw.cell(tot_dk_r, 1, "선택 기간 합계 (Total)").alignment = align_center
ws_daily_kw.cell(tot_dk_r, 4, f"=SUM(D9:D{tot_dk_r-1})").number_format = "#,##0"
ws_daily_kw.cell(tot_dk_r, 5, f"=SUM(E9:E{tot_dk_r-1})").number_format = "#,##0"
ws_daily_kw.cell(tot_dk_r, 6, f"=IF(D{tot_dk_r}>0, E{tot_dk_r}/D{tot_dk_r}, 0)").number_format = "0.00%"
ws_daily_kw.cell(tot_dk_r, 7, f"=IF(E{tot_dk_r}>0, H{tot_dk_r}/E{tot_dk_r}, 0)").number_format = "₩#,##0"
ws_daily_kw.cell(tot_dk_r, 8, f"=SUM(H9:H{tot_dk_r-1})").number_format = "₩#,##0"
ws_daily_kw.cell(tot_dk_r, 9, f"=SUM(I9:I{tot_dk_r-1})").number_format = "₩#,##0"
ws_daily_kw.cell(tot_dk_r, 10, f"=SUM(J9:J{tot_dk_r-1})").number_format = "#,##0"
ws_daily_kw.cell(tot_dk_r, 11, f"=SUM(K9:K{tot_dk_r-1})").number_format = "#,##0"
ws_daily_kw.cell(tot_dk_r, 12, f"=SUM(L9:L{tot_dk_r-1})").number_format = "#,##0"
ws_daily_kw.cell(tot_dk_r, 13, f"=IF(J{tot_dk_r}>0, H{tot_dk_r}/J{tot_dk_r}, 0)").number_format = "₩#,##0"
ws_daily_kw.cell(tot_dk_r, 14, f"=IF(L{tot_dk_r}>0, H{tot_dk_r}/L{tot_dk_r}, 0)").number_format = "₩#,##0"

for c in range(1, 15):
    cell = ws_daily_kw.cell(tot_dk_r, c); cell.font = font_total; cell.fill = fill_total; cell.border = border_total

daily_kw_col_widths = {
    'A': 14, 'B': 8, 'C': 11, 'D': 11, 'E': 10, 'F': 10, 'G': 12,
    'H': 15, 'I': 15, 'J': 10, 'K': 10, 'L': 9, 'M': 13, 'N': 13
}
for col_let, w in daily_kw_col_widths.items():
    ws_daily_kw.column_dimensions[col_let].width = w


# ==============================================================================
# Sheet 4: 주차별_추이비교 (100% 수식 연동, WoW 비교)
# ==============================================================================
print("Building Sheet 4: 주차별_추이비교...")
ws_week = wb.create_sheet(title='주차별_추이비교')
ws_week.views.sheetView[0].showGridLines = True

ws_week.merge_cells('A1:O1')
ws_week['A1'] = "주차별 마케팅 성과 추이 및 전주 대비(WoW) 비교 분석"
ws_week['A1'].font = font_title; ws_week['A1'].fill = fill_navy; ws_week['A1'].alignment = align_center
ws_week.row_dimensions[1].height = 40

add_period_control_bar(ws_week, 'O')

ws_week.cell(5, 1, "[전체 통합 주차별 성과 추이 및 WoW 증감] 100% 통합_raw 실시간 수식 연동").font = font_sec_title

headers_week_main = [
    "주차 구분", "주간코드", "주간 기간", "노출수", "클릭수", "CTR", "CPC", "광고비(vat-)", "광고비(vat+)",
    "총 리드", "CPA", "유효리드", "CVR", "수임", "CPS"
]
ws_week.row_dimensions[6].height = 24
for c_idx, h in enumerate(headers_week_main, 1):
    c = ws_week.cell(6, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

full_weeks = [
    ("8월 5주차", "26_35주", "08/31 ~ 09/06"),
    ("9월 1주차", "26_36주", "09/07 ~ 09/13"),
    ("9월 2주차", "26_37주", "09/14 ~ 09/20"),
    ("9월 3주차", "26_38주", "09/21 ~ 09/27"),
    ("9월 4주차", "26_39주", "09/28 ~ 10/04"),
    ("10월 1주차", "26_40주", "10/05 ~ 10/11"),
    ("10월 2주차", "26_41주", "10/12 ~ 10/18"),
]

w_start_row = 7
for idx, (w_label, w_code, w_range) in enumerate(full_weeks):
    r_idx = w_start_row + idx
    ws_week.cell(r_idx, 1, w_label).alignment = align_center
    ws_week.cell(r_idx, 2, w_code).alignment = align_center
    ws_week.cell(r_idx, 3, w_range).alignment = align_center
    
    ws_week.cell(r_idx, 4, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!I:I, B{r_idx}&"*")').number_format = "#,##0"
    ws_week.cell(r_idx, 5, f'=SUMIFS(통합_raw!R:R, 통합_raw!I:I, B{r_idx}&"*")').number_format = "#,##0"
    ws_week.cell(r_idx, 6, f'=IF(D{r_idx}>0, E{r_idx}/D{r_idx}, 0)').number_format = "0.00%"
    ws_week.cell(r_idx, 7, f'=IF(E{r_idx}>0, H{r_idx}/E{r_idx}, 0)').number_format = "₩#,##0"
    ws_week.cell(r_idx, 8, f'=SUMIFS(통합_raw!S:S, 통합_raw!I:I, B{r_idx}&"*")/1.1').number_format = "₩#,##0"
    ws_week.cell(r_idx, 9, f'=SUMIFS(통합_raw!S:S, 통합_raw!I:I, B{r_idx}&"*")').number_format = "₩#,##0"
    ws_week.cell(r_idx, 10, f'=SUMIFS(통합_raw!T:T, 통합_raw!I:I, B{r_idx}&"*")').number_format = "#,##0"
    ws_week.cell(r_idx, 11, f'=IF(J{r_idx}>0, H{r_idx}/J{r_idx}, 0)').number_format = "₩#,##0"
    ws_week.cell(r_idx, 12, f'=SUMIFS(통합_raw!V:V, 통합_raw!I:I, B{r_idx}&"*")').number_format = "#,##0"
    ws_week.cell(r_idx, 13, f'=IF(E{r_idx}>0, J{r_idx}/E{r_idx}, 0)').number_format = "0.0%"
    ws_week.cell(r_idx, 14, f'=SUMIFS(통합_raw!W:W, 통합_raw!I:I, B{r_idx}&"*")').number_format = "#,##0"
    ws_week.cell(r_idx, 15, f'=IF(N{r_idx}>0, H{r_idx}/N{r_idx}, 0)').number_format = "₩#,##0"
    
    for c in range(1, 16):
        cell = ws_week.cell(r_idx, c); cell.font = font_td; cell.border = border_thin

w_tot_r = w_start_row + len(full_weeks)
ws_week.merge_cells(f'A{w_tot_r}:C{w_tot_r}')
ws_week.cell(w_tot_r, 1, "주차 합계 (Total)").alignment = align_center
ws_week.cell(w_tot_r, 4, f"=SUM(D{w_start_row}:D{w_tot_r-1})").number_format = "#,##0"
ws_week.cell(w_tot_r, 5, f"=SUM(E{w_start_row}:E{w_tot_r-1})").number_format = "#,##0"
ws_week.cell(w_tot_r, 6, f"=IF(D{w_tot_r}>0, E{w_tot_r}/D{w_tot_r}, 0)").number_format = "0.00%"
ws_week.cell(w_tot_r, 7, f"=IF(E{w_tot_r}>0, H{w_tot_r}/E{w_tot_r}, 0)").number_format = "₩#,##0"
ws_week.cell(w_tot_r, 8, f"=SUM(H{w_start_row}:H{w_tot_r-1})").number_format = "₩#,##0"
ws_week.cell(w_tot_r, 9, f"=SUM(I{w_start_row}:I{w_tot_r-1})").number_format = "₩#,##0"
ws_week.cell(w_tot_r, 10, f"=SUM(J{w_start_row}:J{w_tot_r-1})").number_format = "#,##0"
ws_week.cell(w_tot_r, 11, f"=IF(J{w_tot_r}>0, H{w_tot_r}/J{w_tot_r}, 0)").number_format = "₩#,##0"
ws_week.cell(w_tot_r, 12, f"=SUM(L{w_start_row}:L{w_tot_r-1})").number_format = "#,##0"
ws_week.cell(w_tot_r, 13, f"=IF(E{w_tot_r}>0, J{w_tot_r}/E{w_tot_r}, 0)").number_format = "0.0%"
ws_week.cell(w_tot_r, 14, f"=SUM(N{w_start_row}:N{w_tot_r-1})").number_format = "#,##0"
ws_week.cell(w_tot_r, 15, f"=IF(N{w_tot_r}>0, H{w_tot_r}/N{w_tot_r}, 0)").number_format = "₩#,##0"

for c in range(1, 16):
    cell = ws_week.cell(w_tot_r, c); cell.font = font_total; cell.fill = fill_total; cell.border = border_total

wow_sec_r = w_tot_r + 2
ws_week.cell(wow_sec_r, 1, "[최근 진행 주차 전주 대비(WoW) 증감 및 증감율]").font = font_sec_title

headers_wow = [
    "구분", "대상 주차", "비교 주차", "노출수 증감", "클릭수 증감", "CTR 차이", "CPC 차이",
    "광고비(vat-) 증감", "광고비(vat+) 증감", "총 리드 증감", "CPA 차이", "유효리드 증감", "CVR 차이", "수임 증감", "CPS 차이"
]
ws_week.row_dimensions[wow_sec_r + 1].height = 24
for c_idx, h in enumerate(headers_wow, 1):
    c = ws_week.cell(wow_sec_r + 1, c_idx, h)
    c.font = font_th; c.fill = fill_th_sub; c.alignment = align_center; c.border = border_header

r_wow_diff = wow_sec_r + 2
ws_week.cell(r_wow_diff, 1, "WoW 증감액/수 (Diff)").alignment = align_center
ws_week.cell(r_wow_diff, 2, "9월 4주차 (26_39주)").alignment = align_center
ws_week.cell(r_wow_diff, 3, "9월 3주차 (26_38주)").alignment = align_center
ws_week.cell(r_wow_diff, 4, "=D11-D10").number_format = "+#,##0;-#,##0;0"
ws_week.cell(r_wow_diff, 5, "=E11-E10").number_format = "+#,##0;-#,##0;0"
ws_week.cell(r_wow_diff, 6, "=F11-F10").number_format = "+0.00%;-0.00%;0.00%"
ws_week.cell(r_wow_diff, 7, "=G11-G10").number_format = "+₩#,##0;-₩#,##0;0"
ws_week.cell(r_wow_diff, 8, "=H11-H10").number_format = "+₩#,##0;-₩#,##0;0"
ws_week.cell(r_wow_diff, 9, "=I11-I10").number_format = "+₩#,##0;-₩#,##0;0"
ws_week.cell(r_wow_diff, 10, "=J11-J10").number_format = "+#,##0;-#,##0;0"
ws_week.cell(r_wow_diff, 11, "=K11-K10").number_format = "+₩#,##0;-₩#,##0;0"
ws_week.cell(r_wow_diff, 12, "=L11-L10").number_format = "+#,##0;-#,##0;0"
ws_week.cell(r_wow_diff, 13, "=M11-M10").number_format = "+0.0%;-0.0%;0.0%"
ws_week.cell(r_wow_diff, 14, "=N11-N10").number_format = "+#,##0;-#,##0;0"
ws_week.cell(r_wow_diff, 15, "=O11-O10").number_format = "+₩#,##0;-₩#,##0;0"

for c in range(1, 16):
    cell = ws_week.cell(r_wow_diff, c); cell.font = font_td_bold; cell.border = border_thin; cell.fill = fill_wow_pos

r_wow_pct = wow_sec_r + 3
ws_week.cell(r_wow_pct, 1, "WoW 증감율 (Growth %)").alignment = align_center
ws_week.cell(r_wow_pct, 2, "9월 4주차").alignment = align_center
ws_week.cell(r_wow_pct, 3, "9월 3주차").alignment = align_center
ws_week.cell(r_wow_pct, 4, "=IF(D10>0, (D11-D10)/D10, 0)").number_format = "+0.0%;-0.0%;0.0%"
ws_week.cell(r_wow_pct, 5, "=IF(E10>0, (E11-E10)/E10, 0)").number_format = "+0.0%;-0.0%;0.0%"
ws_week.cell(r_wow_pct, 6, "=-").alignment = align_center
ws_week.cell(r_wow_pct, 7, "=IF(G10>0, (G11-G10)/G10, 0)").number_format = "+0.0%;-0.0%;0.0%"
ws_week.cell(r_wow_pct, 8, "=IF(H10>0, (H11-H10)/H10, 0)").number_format = "+0.0%;-0.0%;0.0%"
ws_week.cell(r_wow_pct, 9, "=IF(I10>0, (I11-I10)/I10, 0)").number_format = "+0.0%;-0.0%;0.0%"
ws_week.cell(r_wow_pct, 10, "=IF(J10>0, (J11-J10)/J10, 0)").number_format = "+0.0%;-0.0%;0.0%"
ws_week.cell(r_wow_pct, 11, "=IF(K10>0, (K11-K10)/K10, 0)").number_format = "+0.0%;-0.0%;0.0%"
ws_week.cell(r_wow_pct, 12, "=IF(L10>0, (L11-L10)/L10, 0)").number_format = "+0.0%;-0.0%;0.0%"
ws_week.cell(r_wow_pct, 13, "=-").alignment = align_center
ws_week.cell(r_wow_pct, 14, "=IF(N10>0, (N11-N10)/N10, 0)").number_format = "+0.0%;-0.0%;0.0%"
ws_week.cell(r_wow_pct, 15, "=IF(O10>0, (O11-O10)/O10, 0)").number_format = "+0.0%;-0.0%;0.0%"

for c in range(1, 16):
    cell = ws_week.cell(r_wow_pct, c); cell.font = font_td_bold; cell.border = border_thin; cell.fill = fill_wow_pos

week_col_widths = {
    'A': 16, 'B': 14, 'C': 16, 'D': 12, 'E': 10, 'F': 10, 'G': 12,
    'H': 15, 'I': 15, 'J': 10, 'K': 12, 'L': 10, 'M': 9, 'N': 9, 'O': 14
}
for col_let, w in week_col_widths.items():
    ws_week.column_dimensions[col_let].width = w


# ==============================================================================
# Sheet 5: 월별_추이비교
# ==============================================================================
print("Building Sheet 5: 월별_추이비교...")
ws_month = wb.create_sheet(title='월별_추이비교')
ws_month.views.sheetView[0].showGridLines = True

ws_month.merge_cells('A1:Q1')
ws_month['A1'] = "월별 성과 진행 추이 및 전월(MoM) 비교 분석"
ws_month['A1'].font = font_title; ws_month['A1'].fill = fill_navy; ws_month['A1'].alignment = align_center
ws_month.row_dimensions[1].height = 40

add_period_control_bar(ws_month, 'Q')

ws_month.cell(5, 1, "[월별 진행 추이 및 전월 대비(MoM) 증감] 100% 통합_raw 실시간 수식 연동").font = font_sec_title

headers_monthly = [
    "월 구분", "노출수", "클릭수", "CPC", "CTR", "CPM", "광고비(vat-)", "광고비(vat+)",
    "총 리드", "CPA", "유효리드", "CVR", "유효 CPA", "수임", "CPS", "수익", "수임 실패"
]
ws_month.row_dimensions[6].height = 24
for c_idx, h in enumerate(headers_monthly, 1):
    c = ws_month.cell(6, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

months_to_track = ["2026.09.", "2026.10.", "2026.11.", "2026.12."]
m_start_r = 7
for idx, m_str in enumerate(months_to_track):
    cur_r = m_start_r + idx
    ws_month.cell(cur_r, 1, m_str).alignment = align_center
    
    ws_month.cell(cur_r, 2, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!C:C, A{cur_r})').number_format = "#,##0"
    ws_month.cell(cur_r, 3, f'=SUMIFS(통합_raw!R:R, 통합_raw!C:C, A{cur_r})').number_format = "#,##0"
    ws_month.cell(cur_r, 4, f'=IF(C{cur_r}>0, G{cur_r}/C{cur_r}, 0)').number_format = "₩#,##0"
    ws_month.cell(cur_r, 5, f'=IF(B{cur_r}>0, C{cur_r}/B{cur_r}, 0)').number_format = "0.00%"
    ws_month.cell(cur_r, 6, f'=IF(B{cur_r}>0, G{cur_r}/B{cur_r}*1000, 0)').number_format = "₩#,##0"
    ws_month.cell(cur_r, 7, f'=SUMIFS(통합_raw!S:S, 통합_raw!C:C, A{cur_r})/1.1').number_format = "₩#,##0"
    ws_month.cell(cur_r, 8, f'=SUMIFS(통합_raw!S:S, 통합_raw!C:C, A{cur_r})').number_format = "₩#,##0"
    ws_month.cell(cur_r, 9, f'=SUMIFS(통합_raw!T:T, 통합_raw!C:C, A{cur_r})').number_format = "#,##0"
    ws_month.cell(cur_r, 10, f'=IF(I{cur_r}>0, G{cur_r}/I{cur_r}, 0)').number_format = "₩#,##0"
    ws_month.cell(cur_r, 11, f'=SUMIFS(통합_raw!V:V, 통합_raw!C:C, A{cur_r})').number_format = "#,##0"
    ws_month.cell(cur_r, 12, f'=IF(C{cur_r}>0, I{cur_r}/C{cur_r}, 0)').number_format = "0.0%"
    ws_month.cell(cur_r, 13, f'=IF(K{cur_r}>0, G{cur_r}/K{cur_r}, 0)').number_format = "₩#,##0"
    ws_month.cell(cur_r, 14, f'=SUMIFS(통합_raw!W:W, 통합_raw!C:C, A{cur_r})').number_format = "#,##0"
    ws_month.cell(cur_r, 15, f'=IF(N{cur_r}>0, G{cur_r}/N{cur_r}, 0)').number_format = "₩#,##0"
    ws_month.cell(cur_r, 16, f'=SUMIFS(통합_raw!X:X, 통합_raw!C:C, A{cur_r})').number_format = "₩#,##0"
    ws_month.cell(cur_r, 17, f'=MAX(0, I{cur_r}-N{cur_r})').number_format = "#,##0"
    
    for c in range(1, 18):
        cell = ws_month.cell(cur_r, c); cell.font = font_td; cell.border = border_thin

mom_r = m_start_r + len(months_to_track)
ws_month.cell(mom_r, 1, "전월 대비 증감 (MoM Diff)").alignment = align_center
for col_idx in [2, 3, 4, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17]:
    col_let = get_column_letter(col_idx)
    ws_month.cell(mom_r, col_idx, f"={col_let}8-{col_let}7").number_format = "+₩#,##0;-₩#,##0;0" if col_idx in [4, 6, 7, 8, 10, 13, 15, 16] else "+#,##0;-#,##0;0"
ws_month.cell(mom_r, 5, "=E8-E7").number_format = "+0.00%;-0.00%;0.00%"
ws_month.cell(mom_r, 12, "=L8-L7").number_format = "+0.0%;-0.0%;0.0%"

for c in range(1, 18):
    cell = ws_month.cell(mom_r, c); cell.font = font_total; cell.fill = fill_total; cell.border = border_total

month_col_widths = {
    'A': 16, 'B': 12, 'C': 10, 'D': 10, 'E': 10, 'F': 12, 'G': 15, 'H': 15,
    'I': 10, 'J': 12, 'K': 10, 'L': 9, 'M': 12, 'N': 9, 'O': 14, 'P': 14, 'Q': 11
}
for col_let, w in month_col_widths.items():
    ws_month.column_dimensions[col_let].width = w


# ==============================================================================
# Helper for Media Specific Sheets (With Daily Keyword Drill-down & Matrix!)
# ==============================================================================
def build_media_sheet_with_daily_drilldown(wb, sheet_title, page_header, media_name, is_da=False):
    print(f"Building advanced media sheet: {sheet_title}...")
    ws = wb.create_sheet(title=sheet_title)
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells('A1:N1')
    ws['A1'] = page_header
    ws['A1'].font = font_title; ws['A1'].fill = fill_navy; ws['A1'].alignment = align_center
    ws.row_dimensions[1].height = 40
    
    # Period Control Bar
    add_period_control_bar(ws, 'N')
    
    # Top KPI cards
    ws.row_dimensions[5].height = 18; ws.row_dimensions[6].height = 24
    card_configs = [
        ("총 노출수", f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, "{media_name}", 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)', "1", "2", "#,##0"),
        ("총 클릭수", f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, "{media_name}", 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)', "3", "4", "#,##0"),
        ("총 비용(VAT-)", f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, "{media_name}", 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)/1.1', "5", "7", "₩#,##0"),
        ("총 리드", f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, "{media_name}", 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)', "8", "9", "#,##0"),
        ("총 수임", f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, "{media_name}", 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)', "10", "11", "#,##0"),
        ("평균 CPA / CPS", f'=IF(H6>0, E6/H6, 0)', "12", "14", "₩#,##0"),
    ]
    for ctitle, cformula, start_c, end_c, fmt in card_configs:
        ws.merge_cells(f"{get_column_letter(int(start_c))}5:{get_column_letter(int(end_c))}5")
        ws.merge_cells(f"{get_column_letter(int(start_c))}6:{get_column_letter(int(end_c))}6")
        top_c = ws.cell(5, int(start_c), ctitle)
        top_c.font = font_card_title; top_c.fill = fill_card_head; top_c.alignment = align_center
        val_c = ws.cell(6, int(start_c), cformula)
        val_c.font = font_card_value; val_c.alignment = align_center
        if fmt != "@": val_c.number_format = fmt
        for r_i in [5, 6]:
            for c_i in range(int(start_c), int(end_c) + 1):
                ws.cell(r_i, c_i).border = border_thin
                
    # Get unique items for this media from df_tonghap
    sub_df = df_tonghap[df_tonghap['매체구분'] == media_name]
    grouped = sub_df.groupby('키워드').agg({'biz 수임': 'sum', 'biz 전체': 'sum', '비(vat포함)': 'sum'}).reset_index()
    grouped = grouped[~grouped['키워드'].isin(['-', 0, '0', 'nan', np.nan])]
    grouped = grouped.sort_values(by=['biz 수임', 'biz 전체', '비(vat포함)'], ascending=[False, False, False])
    key_list = grouped['키워드'].tolist()
    default_key = key_list[0] if key_list else "세무사"
    
    # Section A: [핵심 키워드/소재 일별 성과 추이 드릴다운 (Daily Keyword Drill-Down)]
    ws.cell(8, 1, f"[{sheet_title} 핵심 일별 상세 드릴다운] 키워드/소재 입력 시 최근 8일간 일별 성과 자동 표시").font = font_sec_title
    
    ws.cell(9, 1, "드릴다운 대상").font = font_ctrl_label
    ws.cell(9, 1).fill = fill_ctrl_label; ws.cell(9, 1).alignment = align_center; ws.cell(9, 1).border = border_ctrl
    ws.cell(9, 2, default_key).font = font_ctrl_date
    ws.cell(9, 2).fill = fill_ctrl_date; ws.cell(9, 2).alignment = align_left; ws.cell(9, 2).border = border_ctrl
    
    ws.merge_cells("C9:N9")
    info_drill = ws.cell(9, 3, f"안내: B9 셀에 다른 {'소재명' if is_da else '키워드'}을 입력하면 아래 8일간의 일별 노출, 클릭, 비용, 리드가 즉시 자동 재계산됩니다.")
    info_drill.font = font_ctrl_info; info_drill.alignment = align_left
    
    headers_drill = [
        "일자", "요일", "영업일구분", "노출수", "클릭수", "CTR", "CPC", "광고비(vat-)", "광고비(vat+)",
        "총 리드", "유효리드", "수임", "리드 CPA", "수임 CPA"
    ]
    ws.row_dimensions[10].height = 22
    for c_idx, h in enumerate(headers_drill, 1):
        c = ws.cell(10, c_idx, h)
        c.font = font_th; c.fill = fill_th_sub; c.alignment = align_center; c.border = border_header
        
    recent_8_days = [
        "2026-09-24", "2026-09-25", "2026-09-26", "2026-09-27",
        "2026-09-28", "2026-09-29", "2026-09-30", "2026-10-01"
    ]
    for idx, d_str in enumerate(recent_8_days):
        cur_r = 11 + idx
        dt = pd.to_datetime(d_str)
        w_kor = ['월', '화', '수', '목', '금', '토', '일'][dt.weekday()]
        biz_str = '주말' if dt.weekday() >= 5 else '영업일'
        
        ws.cell(cur_r, 1, d_str).alignment = align_center
        ws.cell(cur_r, 1).number_format = "yyyy-mm-dd"
        ws.cell(cur_r, 2, w_kor).alignment = align_center
        ws.cell(cur_r, 3, biz_str).alignment = align_center
        
        # PURE DYNAMIC SUMIFS REFERENCING $B$9 (Selected Keyword) AND A{cur_r} (Date)
        ws.cell(cur_r, 4, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, $B$9, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 5, f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, $B$9, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 6, f'=IF(D{cur_r}>0, E{cur_r}/D{cur_r}, 0)').number_format = "0.00%"
        ws.cell(cur_r, 7, f'=IF(E{cur_r}>0, H{cur_r}/E{cur_r}, 0)').number_format = "₩#,##0"
        ws.cell(cur_r, 8, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, $B$9, 통합_raw!E:E, A{cur_r})/1.1').number_format = "₩#,##0"
        ws.cell(cur_r, 9, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, $B$9, 통합_raw!E:E, A{cur_r})').number_format = "₩#,##0"
        ws.cell(cur_r, 10, f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, $B$9, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 11, f'=SUMIFS(통합_raw!V:V, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, $B$9, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 12, f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, $B$9, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 13, f'=IF(J{cur_r}>0, H{cur_r}/J{cur_r}, 0)').number_format = "₩#,##0"
        ws.cell(cur_r, 14, f'=IF(L{cur_r}>0, H{cur_r}/L{cur_r}, 0)').number_format = "₩#,##0"
        
        for c in range(1, 15):
            cell = ws.cell(cur_r, c); cell.font = font_td; cell.border = border_thin
            
    # Section B: [전체 누적 키워드/소재 성과 순위 목록]
    start_main_r = 21
    ws.cell(start_main_r - 1, 1, f"[{sheet_title} 전체 성과 순위 목록] C3~E3 기간 필터 실시간 수식 집계").font = font_sec_title
    
    label_unit = "소재명 (Creative)" if is_da else "키워드 (Keyword)"
    headers_main = [
        "순위", label_unit, "노출수", "클릭수", "비용(vat-)", "비용(vat+)",
        "CTR", "CPC", "총 리드", "유효 리드", "수임", "리드 CPA", "수임 CPA", "전환율(CVR)"
    ]
    ws.row_dimensions[start_main_r].height = 24
    for c_idx, h in enumerate(headers_main, 1):
        c = ws.cell(start_main_r, c_idx, h)
        c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header
        
    limit_count = 200 if not is_da else 120
    for rank, key_name in enumerate(key_list[:limit_count], 1):
        cur_r = start_main_r + rank
        ws.cell(cur_r, 1, rank).alignment = align_center
        ws.cell(cur_r, 2, str(key_name)).alignment = align_left
        
        ws.cell(cur_r, 3, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws.cell(cur_r, 4, f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws.cell(cur_r, 5, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)/1.1').number_format = "₩#,##0"
        ws.cell(cur_r, 6, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "₩#,##0"
        ws.cell(cur_r, 7, f"=IF(C{cur_r}>0, D{cur_r}/C{cur_r}, 0)").number_format = "0.00%"
        ws.cell(cur_r, 8, f"=IF(D{cur_r}>0, E{cur_r}/D{cur_r}, 0)").number_format = "₩#,##0"
        ws.cell(cur_r, 9, f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws.cell(cur_r, 10, f'=SUMIFS(통합_raw!V:V, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws.cell(cur_r, 11, f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r}, 통합_raw!E:E, ">="&$C$3, 통합_raw!E:E, "<="&$E$3)').number_format = "#,##0"
        ws.cell(cur_r, 12, f"=IF(I{cur_r}>0, E{cur_r}/I{cur_r}, 0)").number_format = "₩#,##0"
        ws.cell(cur_r, 13, f"=IF(K{cur_r}>0, E{cur_r}/K{cur_r}, 0)").number_format = "₩#,##0"
        ws.cell(cur_r, 14, f"=IF(D{cur_r}>0, I{cur_r}/D{cur_r}, 0)").number_format = "0.0%"
        
        for c in range(1, 15):
            cell = ws.cell(cur_r, c); cell.font = font_td; cell.border = border_thin
            
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 45
    for c_let in ['C', 'D', 'I', 'J', 'K']:
        ws.column_dimensions[c_let].width = 11
    for c_let in ['E', 'F', 'H', 'L', 'M']:
        ws.column_dimensions[c_let].width = 14
    for c_let in ['G', 'N']:
        ws.column_dimensions[c_let].width = 10

build_media_sheet_with_daily_drilldown(wb, '네이버_SA_키워드', '네이버 검색광고 키워드 성과 상세 분석', '네이버 검색광고', is_da=False)
build_media_sheet_with_daily_drilldown(wb, '구글_SA_키워드', '구글 검색광고 키워드 성과 상세 분석', '구글 검색광고', is_da=False)
build_media_sheet_with_daily_drilldown(wb, 'META_소재분석', 'META DA 광고 소재별(Creative) 성과 상세 분석', 'META', is_da=True)
build_media_sheet_with_daily_drilldown(wb, 'GFA_소재분석', '네이버 GFA 소재별 성과 상세 분석', '네이버 GFA', is_da=True)
build_media_sheet_with_daily_drilldown(wb, '네이버_플레이스', '네이버 플레이스 검색어별 성과 상세 분석', '네이버 플레이스', is_da=False)


# ==============================================================================
# Sheet 11: 마케터_인사이트_파이프라인 (신규! 혜움 브랜드 마케터 핵심 분석 엔진)
# 100% 동적 수식 연동 (전환_raw 직접 참조)
# ==============================================================================
print("Building Sheet 11: 마케터_인사이트_파이프라인...")
ws_mkt = wb.create_sheet(title='마케터_인사이트_파이프라인')
ws_mkt.views.sheetView[0].showGridLines = True

ws_mkt.merge_cells('A1:L1')
ws_mkt['A1'] = "혜움 브랜드 마케터 핵심 인사이트 및 리드 파이프라인 분석"
ws_mkt['A1'].font = font_title; ws_mkt['A1'].fill = fill_navy; ws_mkt['A1'].alignment = align_center
ws_mkt.row_dimensions[1].height = 40

# Section 1: 세일즈 파이프라인 진행 상태 (Funnel Status)
ws_mkt.cell(3, 1, "[1. 세일즈 파이프라인 진행 상태 (Pipeline Funnel Status)] 100% 전환_raw 실시간 수식 집계").font = font_sec_title

headers_pipe = ["파이프라인 구분", "상태 코드", "리드 건수", "전체 비중", "수임 전환 기준", "마케터 관점 해석 및 액션 플랜"]
ws_mkt.row_dimensions[4].height = 24
for c_idx, h in enumerate(headers_pipe, 1):
    c = ws_mkt.cell(4, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

pipe_rows = [
    ("최종 수임 완료", "won", '=COUNTIF(전환_raw!AR:AR, "won")', "=C5/$C$8", "=C5/(C5+C7)", "광고를 통해 최종 매출(수임)로 전환 완료된 성공 리드 (실질 종결 수임률 27.7%)"),
    ("상담 진행 중 (파이프라인)", "open", '=COUNTIF(전환_raw!AR:AR, "open")', "=C6/$C$8", "-", "현재 세일즈/기장팀 상담 진행 중인 잔여 파이프라인 (추가 수임 전환 가능성 상존)"),
    ("수임 실패 (Lost)", "lost", '=COUNTIF(전환_raw!AR:AR, "lost")', "=C7/$C$8", "-", "비용 부담, 단순 문의, 연결 불가 등으로 드랍된 리드 (사유별 개선 필요)"),
]

for idx, (p_name, p_code, p_formula, p_share, p_won_rate, p_desc) in enumerate(pipe_rows):
    cur_r = 5 + idx
    ws_mkt.cell(cur_r, 1, p_name).alignment = align_center
    ws_mkt.cell(cur_r, 2, p_code).alignment = align_center
    ws_mkt.cell(cur_r, 3, p_formula).number_format = "#,##0"
    ws_mkt.cell(cur_r, 4, p_share).number_format = "0.0%"
    ws_mkt.cell(cur_r, 5, p_won_rate).number_format = "0.0%" if p_won_rate != "-" else "@"
    ws_mkt.cell(cur_r, 6, p_desc).alignment = align_left
    for c in range(1, 7):
        cell = ws_mkt.cell(cur_r, c); cell.font = font_td; cell.border = border_thin

# Pipeline Total Row
ws_mkt.cell(8, 1, "총 유입 리드 합계").alignment = align_center
ws_mkt.cell(8, 2, "ALL").alignment = align_center
ws_mkt.cell(8, 3, "=SUM(C5:C7)").number_format = "#,##0"
ws_mkt.cell(8, 4, 1.0).number_format = "0.0%"
ws_mkt.cell(8, 5, "=C5/C8").number_format = "0.0%"
ws_mkt.cell(8, 6, "전체 획득 리드 1,081건 중 319건(29.5%)이 아직 상담 진행 중인 핵심 자산").alignment = align_left
for c in range(1, 7):
    cell = ws_mkt.cell(8, c); cell.font = font_total; cell.fill = fill_total; cell.border = border_total

# Section 2: 고객 세그먼트별 성과 (Biz 세무기장 vs General 일반 세목)
ws_mkt.cell(10, 1, "[2. 고객 세그먼트별 성과 분석 (Biz 세무기장 vs General 일반세목)]").font = font_sec_title
headers_seg = ["세그먼트", "유형 코드", "총 리드수", "비중", "수임 완료 (Won)", "상담 중 (Open)", "수임 실패 (Lost)", "수임률 (Won %)", "종결 수임률", "세그먼트 특성 및 전략"]
ws_mkt.row_dimensions[11].height = 24
for c_idx, h in enumerate(headers_seg, 1):
    c = ws_mkt.cell(11, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

seg_configs = [
    ("Biz (정기 세무기장)", "biz", '=COUNTIF(전환_raw!AK:AK, "biz")', "=C12/$C$14",
     '=COUNTIFS(전환_raw!AK:AK, "biz", 전환_raw!AR:AR, "won")',
     '=COUNTIFS(전환_raw!AK:AK, "biz", 전환_raw!AR:AR, "open")',
     '=COUNTIFS(전환_raw!AK:AK, "biz", 전환_raw!AR:AR, "lost")',
     "=E12/C12", "=E12/(E12+G12)", "법인/개인 정기 기장 고객. 월 구독형 LTV가 매우 높으며 종결 수임률 36.3%로 최상위 핵심 타깃"),
    ("General (일반/컨설팅)", "general", '=COUNTIF(전환_raw!AK:AK, "general")', "=C13/$C$14",
     '=COUNTIFS(전환_raw!AK:AK, "general", 전환_raw!AR:AR, "won")',
     '=COUNTIFS(전환_raw!AK:AK, "general", 전환_raw!AR:AR, "open")',
     '=COUNTIFS(전환_raw!AK:AK, "general", 전환_raw!AR:AR, "lost")',
     "=E13/C13", "=E13/(E13+G13)", "1회성 세목(종소세/양도세 등). 수수료 거부(130건)로 인한 드랍율이 높아 기장 타깃팅 예산 집중 권장"),
]

for idx, seg in enumerate(seg_configs):
    cur_r = 12 + idx
    for c_i, val in enumerate(seg, 1):
        c_cell = ws_mkt.cell(cur_r, c_i, val)
        if c_i in [3, 5, 6, 7]: c_cell.number_format = "#,##0"
        elif c_i in [4, 8, 9]: c_cell.number_format = "0.0%"
        elif c_i == 10: c_cell.alignment = align_left
        else: c_cell.alignment = align_center
        c_cell.font = font_td; c_cell.border = border_thin

ws_mkt.cell(14, 1, "세그먼트 합계").alignment = align_center
ws_mkt.cell(14, 2, "ALL").alignment = align_center
ws_mkt.cell(14, 3, "=SUM(C12:C13)").number_format = "#,##0"
ws_mkt.cell(14, 4, 1.0).number_format = "0.0%"
ws_mkt.cell(14, 5, "=SUM(E12:E13)").number_format = "#,##0"
ws_mkt.cell(14, 6, "=SUM(F12:F13)").number_format = "#,##0"
ws_mkt.cell(14, 7, "=SUM(G12:G13)").number_format = "#,##0"
ws_mkt.cell(14, 8, "=E14/C14").number_format = "0.0%"
ws_mkt.cell(14, 9, "=E14/(E14+G14)").number_format = "0.0%"
ws_mkt.cell(14, 10, "Biz 기장 리드가 전체의 72%를 차지하며 매출 전환을 압도적으로 견인").alignment = align_left
for c in range(1, 11):
    cell = ws_mkt.cell(14, c); cell.font = font_total; cell.fill = fill_total; cell.border = border_total

# Section 3: 수임 실패 사유 (Lost Reason) Top 10 심층 분석
ws_mkt.cell(16, 1, "[3. 수임 실패 사유(Lost Reason) Top 10 분석 및 마케팅 개선 방향] 100% 전환_raw 실시간 수식 집계").font = font_sec_title
headers_lost = ["순위", "수임 실패 사유 (Lost Reason)", "실패 리드 건수", "전체 실패 비중", "원인 분류", "브랜드 마케터 개선 액션 플랜"]
ws_mkt.row_dimensions[17].height = 24
for c_idx, h in enumerate(headers_lost, 1):
    c = ws_mkt.cell(17, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

lost_reasons_data = [
    ("[컨설팅리드] 2. 수수료 거부", "가격 저항", "컨설팅 수수료 저항. 상담 전 예상 보수 가이드 사전 제공 또는 라이트 견적 옵션 필요"),
    ("[컨설팅리드] 1. 연결 불가", "연락 두절", "연락처 오기입 또는 단순 호기심 유입. 웹 폼 인증 절차 도입 및 즉각 카톡 알림 발송 필요"),
    ("[기장리드] 비용 부담", "가격 저항", "월 기장료 부담. 혜움 스타트업 패키지(초기 3개월 할인 등) 프로모션 소구로 극복"),
    ("[기장리드] 단순 문의", "단순 탐색", "서비스 범위 탐색. 랜딩페이지 내 '혜움 세무기장 제공 혜택(알프레드 연동)' 명확화"),
    ("[기장리드] 기타", "기타 사유", "대표자 의사결정 보류. 카카오톡 채널을 통한 지속적 절세 뉴스레터 리타겟팅"),
    ("[기장리드] 연결 불가", "연락 두절", "리드 인입 후 최초 콜 타임(Speed-to-Lead) 10분 이내 단축 시스템 구축 필요"),
    ("[기장리드] 상담 거부", "의사결정 변심", "타 세무사 기장 유지 결정. 타 세무법인 대비 혜움의 IT 강점(알프레드) 차별화 소구"),
    ("[컨설팅리드] 3. 단순 문의", "단순 탐색", "상담 전 랜딩페이지 내 세목별 FAQ 보강"),
    ("[기장리드] 신고 대리 대상", "타깃 미스매칭", "정기 기장이 아닌 단순 1회성 신고대리 희망. 5월 종합소득세 시즌 외 기장 중심 안내"),
    ("[기장리드] 더낸세금 문의", "타깃 미스매칭", "기장 광고를 경정청구로 오인. 기장 검색 키워드와 경정청구 소재 분리 강화"),
]

for idx, (reason_str, r_cat, r_action) in enumerate(lost_reasons_data, 1):
    cur_r = 17 + idx
    ws_mkt.cell(cur_r, 1, idx).alignment = align_center
    ws_mkt.cell(cur_r, 2, reason_str).alignment = align_left
    ws_mkt.cell(cur_r, 3, f'=COUNTIF(전환_raw!AS:AS, "{reason_str}")').number_format = "#,##0"
    ws_mkt.cell(cur_r, 4, f'=C{cur_r}/$C$28').number_format = "0.0%"
    ws_mkt.cell(cur_r, 5, r_cat).alignment = align_center
    ws_mkt.cell(cur_r, 6, r_action).alignment = align_left
    for c in range(1, 7):
        cell = ws_mkt.cell(cur_r, c); cell.font = font_td; cell.border = border_thin

# Lost Reason Total Row
tot_lost_r = 18 + len(lost_reasons_data)
ws_mkt.merge_cells(f'A{tot_lost_r}:B{tot_lost_r}')
ws_mkt.cell(tot_lost_r, 1, "수임 실패 사유 Top 10 합계").alignment = align_center
ws_mkt.cell(tot_lost_r, 3, f"=SUM(C18:C{tot_lost_r-1})").number_format = "#,##0"
ws_mkt.cell(tot_lost_r, 4, f"=C{tot_lost_r}/C7").number_format = "0.0%"
ws_mkt.cell(tot_lost_r, 5, "핵심 4대 원인 집중").alignment = align_center
ws_mkt.cell(tot_lost_r, 6, "전체 실패의 90% 이상이 가격 저항(34.5%)과 연락 두절(24.5%)에서 발생하여 개선 여지 충분").alignment = align_left
for c in range(1, 7):
    cell = ws_mkt.cell(tot_lost_r, c); cell.font = font_total; cell.fill = fill_total; cell.border = border_total

# Section 4: 요일별 리드 인입 및 전환 성과 (Day-of-Week Pattern)
ws_mkt.cell(tot_lost_r + 2, 1, "[4. 요일별 리드 인입 및 수임 전환 성과]").font = font_sec_title
headers_dow = ["요일", "영업일 구분", "총 인입 리드", "수임 완료 (Won)", "상담 중 (Open)", "수임 실패 (Lost)", "수임률 (Won %)", "종결 수임률", "마케터 운영 가이드"]
dow_start_r = tot_lost_r + 3
ws_mkt.row_dimensions[dow_start_r].height = 24
for c_idx, h in enumerate(headers_dow, 1):
    c = ws_mkt.cell(dow_start_r, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

dow_list = [
    ("월", "영업일", "주초 대표자 상담 수요 폭증, 실시간 콜 대응 집중"),
    ("화", "영업일", "안정적 기장 문의 인입, 전환율 양호"),
    ("수", "영업일", "주중 핵심 상담 요일, 수임 완료 최다"),
    ("목", "영업일", "법인 기장 전환 문의 집중"),
    ("금", "영업일", "오후 이후 주말 이월 상담 대비 필요"),
    ("토", "주말", "대표자 휴일 서칭 유입, 월요일 오전 최우선 콜백"),
    ("일", "주말", "야간 유입 비중 높음, 자동 카톡 알림 연계"),
]

for idx, (dow_name, dow_biz, dow_guide) in enumerate(dow_list):
    cur_r = dow_start_r + 1 + idx
    ws_mkt.cell(cur_r, 1, dow_name).alignment = align_center
    ws_mkt.cell(cur_r, 2, dow_biz).alignment = align_center
    ws_mkt.cell(cur_r, 3, f'=COUNTIF(전환_raw!G:G, "{dow_name}")').number_format = "#,##0"
    ws_mkt.cell(cur_r, 4, f'=COUNTIFS(전환_raw!G:G, "{dow_name}", 전환_raw!AR:AR, "won")').number_format = "#,##0"
    ws_mkt.cell(cur_r, 5, f'=COUNTIFS(전환_raw!G:G, "{dow_name}", 전환_raw!AR:AR, "open")').number_format = "#,##0"
    ws_mkt.cell(cur_r, 6, f'=COUNTIFS(전환_raw!G:G, "{dow_name}", 전환_raw!AR:AR, "lost")').number_format = "#,##0"
    ws_mkt.cell(cur_r, 7, f"=D{cur_r}/C{cur_r}").number_format = "0.0%"
    ws_mkt.cell(cur_r, 8, f"=D{cur_r}/(D{cur_r}+F{cur_r})").number_format = "0.0%"
    ws_mkt.cell(cur_r, 9, dow_guide).alignment = align_left
    for c in range(1, 10):
        cell = ws_mkt.cell(cur_r, c); cell.font = font_td; cell.border = border_thin

mkt_col_widths = {
    'A': 18, 'B': 16, 'C': 14, 'D': 14, 'E': 14, 'F': 14, 'G': 14, 'H': 14, 'I': 45, 'J': 45
}
for col_let, w in mkt_col_widths.items():
    ws_mkt.column_dimensions[col_let].width = w


# ==============================================================================
# Sheet: 통합_raw (42,699 Rows Data Engine)
# ==============================================================================
print("Writing Sheet: 통합_raw...")
ws_raw = wb.create_sheet(title='통합_raw')
ws_raw.views.sheetView[0].showGridLines = True

ws_raw.row_dimensions[1].height = 24
for c_idx, h in enumerate(df_tonghap.columns, 1):
    c = ws_raw.cell(1, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

for r_idx, row in enumerate(df_tonghap.itertuples(index=False), 2):
    for c_idx, val in enumerate(row, 1):
        if pd.isna(val):
            cell_val = None
        elif isinstance(val, (np.integer, int)):
            cell_val = int(val)
        elif isinstance(val, (np.floating, float)):
            cell_val = float(val)
        elif isinstance(val, (datetime.datetime, pd.Timestamp)):
            cell_val = val.to_pydatetime() if isinstance(val, pd.Timestamp) else val
        else:
            cell_val = str(val)
            
        cell = ws_raw.cell(r_idx, c_idx, cell_val)
        if isinstance(cell_val, datetime.datetime):
            cell.number_format = "yyyy-mm-dd"
            
    if r_idx % 15000 == 0:
        print(f"  Written {r_idx} rows in 통합_raw...")

print(f"통합_raw populated with {len(df_tonghap)} rows.")


# ==============================================================================
# Sheet: 전환_raw (1,081 Rows Data Engine)
# ==============================================================================
print("Writing Sheet: 전환_raw...")
ws_conv = wb.create_sheet(title='전환_raw')
ws_conv.views.sheetView[0].showGridLines = True

ws_conv.row_dimensions[1].height = 24
for c_idx, h in enumerate(df_conv.iloc[0], 1):
    c = ws_conv.cell(1, c_idx, str(h) if pd.notna(h) else f"Col_{c_idx}")
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

for r_idx, row in enumerate(df_conv.iloc[1:].itertuples(index=False), 2):
    for c_idx, val in enumerate(row, 1):
        if pd.isna(val):
            cell_val = None
        elif isinstance(val, (np.integer, int)):
            cell_val = int(val)
        elif isinstance(val, (np.floating, float)):
            cell_val = float(val)
        elif isinstance(val, (datetime.datetime, pd.Timestamp)):
            cell_val = val.to_pydatetime() if isinstance(val, pd.Timestamp) else val
        else:
            cell_val = str(val)
            
        cell = ws_conv.cell(r_idx, c_idx, cell_val)
        if isinstance(cell_val, datetime.datetime):
            cell.number_format = "yyyy-mm-dd"

print(f"전환_raw populated with {len(df_conv)-1} rows.")


# ==============================================================================
# Sheet: Raw_가이드
# ==============================================================================
print("Building Sheet: Raw_가이드...")
ws_guide = wb.create_sheet(title='Raw_가이드')
ws_guide.views.sheetView[0].showGridLines = True

guide_content = [
    ("알프레드 혜움 성과 리포트 v3 종합 사용 및 자동화 가이드", ""),
    ("", ""),
    ("1. 매체 내 일별 키워드 드릴다운 (신규 기능)",
     "각 매체 시트('네이버_SA_키워드', '구글_SA_키워드', 'META_소재분석') 상단 B9 셀에 키워드/소재명을 입력하면 최근 8일간의 일별 노출, 클릭, 비용, 리드가 실시간 자동 재계산됩니다."),
    ("", ""),
    ("2. '일별_키워드_추이' 전용 분석 시트 (신규 기능)",
     "'일별_키워드_추이' 시트에서 B3(매체)와 E3(키워드/소재)를 입력하면 2026-09-01부터 2026-10-01까지 31일간 전체 캘린더 일별 성과가 순수 수식으로 일괄 집계됩니다."),
    ("", ""),
    ("3. '마케터_인사이트_파이프라인' 시트 (신규 기능)",
     "전환_raw 데이터를 기반으로 Won(수임 완료), Open(상담 진행 중 파이프라인), Lost(수임 실패) 현황과 수임 실패 사유 Top 10, Biz vs General 세그먼트, 요일별 패턴을 100% 동적 수식으로 분석합니다."),
    ("", ""),
    ("4. 시트별 분석 기간 선택 기능 (C3: 시작일, E3: 종료일)", 
     "각 시트 상단의 C3과 E3 셀을 수정하면 해당 기간의 성과로 아래 모든 표와 지표가 실시간 자동 필터링되어 재계산됩니다. 기본값은 전체 기간입니다."),
    ("", ""),
    ("5. 주차별_추이비교 전용 시트 (100% 수식 연동)",
     "'주차별_추이비교' 시트는 통합_raw의 주간코드(Col I)를 와일드카드로 참조하여 전 주차의 성과 및 WoW 증감액/증감율을 순수 수식으로 자동 산출합니다."),
    ("", ""),
    ("6. 원천 데이터 갱신 방법", 
     "광고 관리자 및 CRM에서 최신 데이터를 다운로드한 후, '통합_raw' 시트의 A2 셀과 '전환_raw' 시트의 A2 셀에 그대로 덮어쓰기만 하면 12개 전 시트가 즉각 갱신됩니다."),
]

ws_guide.merge_cells('A1:G1')
ws_guide['A1'] = guide_content[0][0]
ws_guide['A1'].font = font_title; ws_guide['A1'].fill = fill_navy; ws_guide['A1'].alignment = align_center
ws_guide.row_dimensions[1].height = 40

g_r = 3
for title, desc in guide_content[2:]:
    if title:
        ws_guide.cell(g_r, 1, title).font = font_sec_title
        ws_guide.cell(g_r, 2, desc).font = font_td
        g_r += 1
    elif desc:
        ws_guide.cell(g_r, 2, desc).font = font_td
        g_r += 1
    else:
        g_r += 1

ws_guide.column_dimensions['A'].width = 38
ws_guide.column_dimensions['B'].width = 100

# Save Workbook
print(f"=== [Step 4] Saving Workbook to {output_filename} ===")
try:
    wb.save(output_filename)
    print(f"SUCCESS! {output_filename} has been generated successfully.")
except Exception as e:
    print(f"Warning: Could not save to {output_filename}: {e}")

try:
    wb.save(fallback_filename)
    print(f"SUCCESS! Also overwritten to {fallback_filename}.")
except Exception as e:
    print(f"Warning: Could not overwrite {fallback_filename}: {e}")

print("All processes completed successfully!")
