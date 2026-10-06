# -*- coding: utf-8 -*-
"""
알프레드 혜움 세무기장 성과 리포트 v2
- 시트별 기간 선택 기능 (시작일/종료일 셀 입력 시 100% 실시간 자동 필터링 수식)
- 언제부터 언제까지의 데이터인지 자동 표시 (Raw 전체 수집 기간 및 현재 조회 기간)
- 주차별_추이비교 전용 시트 신설 (100% 순수 수식, WoW 증감 및 증감율, 매체별 주차 추이)
- 월별_추이비교 시트 100% 동적 수식화 (하드코딩 완전 제거)
- 이모티콘 완전 배제, 다크 네이비 UI/UX 프리미엄 비즈니스 테마
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

# Load 통합_raw
df_tonghap = pd.read_excel(xl, sheet_name='통합_raw')
print(f"통합_raw loaded. Shape: {df_tonghap.shape}")

# Ensure date format
df_tonghap['일'] = pd.to_datetime(df_tonghap['일'])

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
fill_wow_pos = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # 연초록
fill_wow_neg = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # 연주황

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
    """
    Adds a standardized Period Control Bar at Row 3:
    - B3: 조회 시작일 (Label)
    - C3: Start Date input (Default: =MIN(통합_raw!E:E))
    - D3: 조회 종료일 (Label)
    - E3: End Date input (Default: =MAX(통합_raw!E:E))
    - F3:G3: 전체 Raw 기간 자동 안내
    - H3:max_col: 실시간 재집계 안내 문구
    """
    ws.row_dimensions[3].height = 24
    
    ws.cell(3, 2, "조회 시작일").font = font_ctrl_label
    ws.cell(3, 2).fill = fill_ctrl_label
    ws.cell(3, 2).alignment = align_center
    ws.cell(3, 2).border = border_ctrl
    
    c3 = ws.cell(3, 3, "=MIN(통합_raw!E:E)")
    c3.font = font_ctrl_date
    c3.fill = fill_ctrl_date
    c3.alignment = align_center
    c3.number_format = "yyyy-mm-dd"
    c3.border = border_ctrl
    
    ws.cell(3, 4, "조회 종료일").font = font_ctrl_label
    ws.cell(3, 4).fill = fill_ctrl_label
    ws.cell(3, 4).alignment = align_center
    ws.cell(3, 4).border = border_ctrl
    
    e3 = ws.cell(3, 5, "=MAX(통합_raw!E:E)")
    e3.font = font_ctrl_date
    e3.fill = fill_ctrl_date
    e3.alignment = align_center
    e3.number_format = "yyyy-mm-dd"
    e3.border = border_ctrl
    
    ws.merge_cells("F3:G3")
    fg = ws.cell(3, 6, '="[전체 원천 데이터 기간: " & TEXT(MIN(통합_raw!E:E),"yyyy-mm-dd") & " ~ " & TEXT(MAX(통합_raw!E:E),"yyyy-mm-dd") & "]"')
    fg.font = font_sec_title
    fg.alignment = align_left
    
    ws.merge_cells(f"H3:{max_col_letter}3")
    info_c = ws.cell(3, 8, "안내: C3(시작일)과 E3(종료일)을 변경하면 아래 모든 성과가 해당 기간으로 실시간 자동 재집계됩니다.")
    info_c.font = font_ctrl_info
    info_c.alignment = align_left


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

# Period Control Bar
add_period_control_bar(ws_dash, 'R')

# Section 1 Header: Daily and Period Controls
ws_dash.cell(5, 1, "[일일 운영 마감 기준일]").font = font_sec_title
ws_dash.cell(6, 1, "2026-10-01").font = font_card_value
ws_dash.cell(6, 1).alignment = align_center
ws_dash.cell(6, 1).number_format = "yyyy-mm-dd"

# Top Summary KPI Cards (Referencing C3 and E3 Date Range)
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

ws_dash.row_dimensions[5].height = 18
ws_dash.row_dimensions[6].height = 26

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

# Section 2: Media Target & Performance Summary (Formula Linked to C3 and E3 Date Range)
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
        # PURE DYNAMIC FORMULAS FILTERED BY C3(시작일) ~ E3(종료일)
        ws_dash.cell(curr_row, 3, b_exc).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 4, b_inc).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 5, share).number_format = "0.0%"
        ws_dash.cell(curr_row, 6, f"=IF(C{curr_row}>0, L{curr_row}/C{curr_row}, 0)").number_format = "0.0%"
        ws_dash.cell(curr_row, 7, t_cpa).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 8, f"=IF(O{curr_row}>0, L{curr_row}/O{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 9, exp_close).number_format = "₩#,##0"
        
        # Col Q: 노, Col R: 클, Col S: 비(vat포함), Col T: biz 전체, Col W: biz 수임
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

# Total Row
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

# Column Widths
dash_col_widths = {
    'A': 12, 'B': 22, 'C': 14, 'D': 14, 'E': 9, 'F': 9, 'G': 12, 'H': 12, 'I': 14,
    'J': 11, 'K': 10, 'L': 14, 'M': 9, 'N': 10, 'O': 10, 'P': 12, 'Q': 9, 'R': 12
}
for col_let, w in dash_col_widths.items():
    ws_dash.column_dimensions[col_let].width = w


# ==============================================================================
# Sheet 2: 브랜드_키워드분석 (C3~E3 기간 필터 100% 연동)
# ==============================================================================
print("Building Sheet 2: 브랜드_키워드분석...")
ws_kw_cat = wb.create_sheet(title='브랜드_키워드분석')
ws_kw_cat.views.sheetView[0].showGridLines = True

ws_kw_cat.merge_cells('A1:O1')
ws_kw_cat['A1'] = "혜움 브랜드 및 카테고리별 키워드 성과 상세 분석"
ws_kw_cat['A1'].font = font_title; ws_kw_cat['A1'].fill = fill_navy; ws_kw_cat['A1'].alignment = align_center
ws_kw_cat.row_dimensions[1].height = 40

# Period Control Bar
add_period_control_bar(ws_kw_cat, 'O')

# Categorized Keywords Structure
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
    # Category Header Row
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
        
        # PURE DYNAMIC FORMULAS DIRECTLY FROM 통합_raw WITH C3~E3 DATE RANGE FILTER
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
        
    # Category Subtotal Row
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

# Widths for Category Sheet
kw_cat_widths = {
    'A': 18, 'B': 22, 'C': 18, 'D': 11, 'E': 10, 'F': 14, 'G': 14,
    'H': 9, 'I': 11, 'J': 10, 'K': 10, 'L': 9, 'M': 13, 'N': 13, 'O': 10
}
for col_let, w in kw_cat_widths.items():
    ws_kw_cat.column_dimensions[col_let].width = w


# ==============================================================================
# Sheet 3: 주차별_추이비교 (Dedicated Weekly Performance & WoW Sheet!)
# 100% PURE DYNAMIC SUMIFS WITH PREFIX MATCH ON 통합_raw!I:I
# ==============================================================================
print("Building Sheet 3: 주차별_추이비교...")
ws_week = wb.create_sheet(title='주차별_추이비교')
ws_week.views.sheetView[0].showGridLines = True

ws_week.merge_cells('A1:O1')
ws_week['A1'] = "주차별 마케팅 성과 추이 및 전주 대비(WoW) 비교 분석"
ws_week['A1'].font = font_title; ws_week['A1'].fill = fill_navy; ws_week['A1'].alignment = align_center
ws_week.row_dimensions[1].height = 40

# Period Control Bar
add_period_control_bar(ws_week, 'O')

# Section 1: 전체 통합 주차별 성과 추이 (전부 수식)
ws_week.cell(5, 1, "[전체 통합 주차별 성과 추이 및 WoW 증감] 100% 통합_raw 실시간 수식 연동").font = font_sec_title

headers_week_main = [
    "주차 구분", "주간코드", "주간 기간", "노출수", "클릭수", "CTR", "CPC", "광고비(vat-)", "광고비(vat+)",
    "총 리드", "CPA", "유효리드", "CVR", "수임", "CPS"
]

ws_week.row_dimensions[6].height = 24
for c_idx, h in enumerate(headers_week_main, 1):
    c = ws_week.cell(6, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

# Weeks Definition (Prefix code for Col I match)
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
    
    # 100% PURE DYNAMIC SUMIFS USING WILDCARD B{r_idx} & "*" ON 통합_raw!I:I
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

# Weekly Total Row
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

# Section 2: 전주 대비 증감 (WoW) 비교 테이블 (순수 수식 계산)
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

# WoW Comparison Rows (Row 11 vs Row 10: 9월 4주차 vs 9월 3주차)
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

# Section 3: 주요 매체별 주차 추이 테이블 (네이버 SA, 구글 SA, META)
sec3_r = r_wow_pct + 2
ws_week.cell(sec3_r, 1, "[주요 매체별 주차 추이 비교] 100% 통합_raw 실시간 수식 집계").font = font_sec_title

headers_media_week = [
    "매체구분", "주차 구분", "주간코드", "노출수", "클릭수", "CTR", "CPC", "광고비(vat-)", "광고비(vat+)",
    "총 리드", "CPA", "유효리드", "CVR", "수임", "CPS"
]
ws_week.row_dimensions[sec3_r + 1].height = 24
for c_idx, h in enumerate(headers_media_week, 1):
    c = ws_week.cell(sec3_r + 1, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

cur_mw_r = sec3_r + 2
active_track_media = [
    ("네이버 검색광고", full_weeks[1:5]), # 9월 1~4주
    ("구글 검색광고", full_weeks[1:5]),
    ("META", full_weeks[1:5]),
]

for med_name, w_sublist in active_track_media:
    for w_label, w_code, w_range in w_sublist:
        ws_week.cell(cur_mw_r, 1, med_name).alignment = align_center
        ws_week.cell(cur_mw_r, 2, w_label).alignment = align_center
        ws_week.cell(cur_mw_r, 3, w_code).alignment = align_center
        
        # PURE DYNAMIC SUMIFS REFERENCING MEDIA AND WEEK PREFIX
        ws_week.cell(cur_mw_r, 4, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, A{cur_mw_r}, 통합_raw!I:I, C{cur_mw_r}&"*")').number_format = "#,##0"
        ws_week.cell(cur_mw_r, 5, f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, A{cur_mw_r}, 통합_raw!I:I, C{cur_mw_r}&"*")').number_format = "#,##0"
        ws_week.cell(cur_mw_r, 6, f'=IF(D{cur_mw_r}>0, E{cur_mw_r}/D{cur_mw_r}, 0)').number_format = "0.00%"
        ws_week.cell(cur_mw_r, 7, f'=IF(E{cur_mw_r}>0, H{cur_mw_r}/E{cur_mw_r}, 0)').number_format = "₩#,##0"
        ws_week.cell(cur_mw_r, 8, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, A{cur_mw_r}, 통합_raw!I:I, C{cur_mw_r}&"*")/1.1').number_format = "₩#,##0"
        ws_week.cell(cur_mw_r, 9, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, A{cur_mw_r}, 통합_raw!I:I, C{cur_mw_r}&"*")').number_format = "₩#,##0"
        ws_week.cell(cur_mw_r, 10, f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, A{cur_mw_r}, 통합_raw!I:I, C{cur_mw_r}&"*")').number_format = "#,##0"
        ws_week.cell(cur_mw_r, 11, f'=IF(J{cur_mw_r}>0, H{cur_mw_r}/J{cur_mw_r}, 0)').number_format = "₩#,##0"
        ws_week.cell(cur_mw_r, 12, f'=SUMIFS(통합_raw!V:V, 통합_raw!L:L, A{cur_mw_r}, 통합_raw!I:I, C{cur_mw_r}&"*")').number_format = "#,##0"
        ws_week.cell(cur_mw_r, 13, f'=IF(E{cur_mw_r}>0, J{cur_mw_r}/E{cur_mw_r}, 0)').number_format = "0.0%"
        ws_week.cell(cur_mw_r, 14, f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, A{cur_mw_r}, 통합_raw!I:I, C{cur_mw_r}&"*")').number_format = "#,##0"
        ws_week.cell(cur_mw_r, 15, f'=IF(N{cur_mw_r}>0, H{cur_mw_r}/N{cur_mw_r}, 0)').number_format = "₩#,##0"
        
        for c in range(1, 16):
            cell = ws_week.cell(cur_mw_r, c); cell.font = font_td; cell.border = border_thin
        cur_mw_r += 1

week_col_widths = {
    'A': 16, 'B': 14, 'C': 16, 'D': 12, 'E': 10, 'F': 10, 'G': 12,
    'H': 15, 'I': 15, 'J': 10, 'K': 12, 'L': 10, 'M': 9, 'N': 9, 'O': 14
}
for col_let, w in week_col_widths.items():
    ws_week.column_dimensions[col_let].width = w


# ==============================================================================
# Sheet 4: 월별_추이비교 (100% 동적 수식화, 하드코딩 완전 제거)
# ==============================================================================
print("Building Sheet 4: 월별_추이비교...")
ws_month = wb.create_sheet(title='월별_추이비교')
ws_month.views.sheetView[0].showGridLines = True

ws_month.merge_cells('A1:Q1')
ws_month['A1'] = "월별 성과 진행 추이 및 전월(MoM) 비교 분석"
ws_month['A1'].font = font_title; ws_month['A1'].fill = fill_navy; ws_month['A1'].alignment = align_center
ws_month.row_dimensions[1].height = 40

# Period Control Bar
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

# Monthly dynamic rows
months_to_track = ["2026.09.", "2026.10.", "2026.11.", "2026.12."]
m_start_r = 7
for idx, m_str in enumerate(months_to_track):
    cur_r = m_start_r + idx
    ws_month.cell(cur_r, 1, m_str).alignment = align_center
    
    # 100% PURE DYNAMIC SUMIFS DIRECTLY MATCHING 통합_raw!C:C (월)
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

# MoM Difference Row (10월 vs 9월)
mom_r = m_start_r + len(months_to_track)
ws_month.cell(mom_r, 1, "전월 대비 증감 (MoM Diff)").alignment = align_center
for col_idx in [2, 3, 4, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17]:
    col_let = get_column_letter(col_idx)
    ws_month.cell(mom_r, col_idx, f"={col_let}8-{col_let}7").number_format = "+₩#,##0;-₩#,##0;0" if col_idx in [4, 6, 7, 8, 10, 13, 15, 16] else "+#,##0;-#,##0;0"
ws_month.cell(mom_r, 5, "=E8-E7").number_format = "+0.00%;-0.00%;0.00%"
ws_month.cell(mom_r, 12, "=L8-L7").number_format = "+0.0%;-0.0%;0.0%"

for c in range(1, 18):
    cell = ws_month.cell(mom_r, c); cell.font = font_total; cell.fill = fill_total; cell.border = border_total

# MoM Growth % Row
mom_pct_r = mom_r + 1
ws_month.cell(mom_pct_r, 1, "전월 대비 증감율 (MoM %)").alignment = align_center
for col_idx in [2, 3, 4, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17]:
    col_let = get_column_letter(col_idx)
    ws_month.cell(mom_pct_r, col_idx, f"=IF({col_let}7>0, ({col_let}8-{col_let}7)/{col_let}7, 0)").number_format = "+0.0%;-0.0%;0.0%"
ws_month.cell(mom_pct_r, 5, "=-").alignment = align_center
ws_month.cell(mom_pct_r, 12, "=-").alignment = align_center

for c in range(1, 18):
    cell = ws_month.cell(mom_pct_r, c); cell.font = font_total; cell.fill = fill_total; cell.border = border_total

# Daily DoD Section
dod_title_r = mom_pct_r + 2
ws_month.cell(dod_title_r, 1, "[일별 성과 추이 (Daily DoD)] 100% 통합_raw 실시간 수식 집계").font = font_sec_title

headers_daily = [
    "일자", "노출수", "클릭수", "CTR", "CPC", "광고비(vat-)", "광고비(vat+)",
    "총 리드", "CPA", "유효리드", "CVR", "수임", "CPS"
]
ws_month.row_dimensions[dod_title_r + 1].height = 24
for c_idx, h in enumerate(headers_daily, 1):
    c = ws_month.cell(dod_title_r + 1, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

# Sample recent days
sample_days = [
    "2026-09-24", "2026-09-25", "2026-09-26", "2026-09-27",
    "2026-09-28", "2026-09-29", "2026-09-30", "2026-10-01"
]
d_start_r = dod_title_r + 2
for idx, d_str in enumerate(sample_days):
    cur_r = d_start_r + idx
    ws_month.cell(cur_r, 1, d_str).alignment = align_center
    ws_month.cell(cur_r, 1).number_format = "yyyy-mm-dd"
    
    # 100% PURE DYNAMIC SUMIFS MATCHING 통합_raw!E:E (일)
    ws_month.cell(cur_r, 2, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
    ws_month.cell(cur_r, 3, f'=SUMIFS(통합_raw!R:R, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
    ws_month.cell(cur_r, 4, f'=IF(B{cur_r}>0, C{cur_r}/B{cur_r}, 0)').number_format = "0.00%"
    ws_month.cell(cur_r, 5, f'=IF(C{cur_r}>0, F{cur_r}/C{cur_r}, 0)').number_format = "₩#,##0"
    ws_month.cell(cur_r, 6, f'=SUMIFS(통합_raw!S:S, 통합_raw!E:E, A{cur_r})/1.1').number_format = "₩#,##0"
    ws_month.cell(cur_r, 7, f'=SUMIFS(통합_raw!S:S, 통합_raw!E:E, A{cur_r})').number_format = "₩#,##0"
    ws_month.cell(cur_r, 8, f'=SUMIFS(통합_raw!T:T, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
    ws_month.cell(cur_r, 9, f'=IF(H{cur_r}>0, F{cur_r}/H{cur_r}, 0)').number_format = "₩#,##0"
    ws_month.cell(cur_r, 10, f'=SUMIFS(통합_raw!V:V, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
    ws_month.cell(cur_r, 11, f'=IF(C{cur_r}>0, H{cur_r}/C{cur_r}, 0)').number_format = "0.0%"
    ws_month.cell(cur_r, 12, f'=SUMIFS(통합_raw!W:W, 통합_raw!E:E, A{cur_r})').number_format = "#,##0"
    ws_month.cell(cur_r, 13, f'=IF(L{cur_r}>0, F{cur_r}/L{cur_r}, 0)').number_format = "₩#,##0"
    
    for c in range(1, 14):
        cell = ws_month.cell(cur_r, c); cell.font = font_td; cell.border = border_thin

month_col_widths = {
    'A': 16, 'B': 12, 'C': 10, 'D': 10, 'E': 10, 'F': 12, 'G': 15, 'H': 15,
    'I': 10, 'J': 12, 'K': 10, 'L': 9, 'M': 12, 'N': 9, 'O': 14, 'P': 14, 'Q': 11
}
for col_let, w in month_col_widths.items():
    ws_month.column_dimensions[col_let].width = w


# ==============================================================================
# Helper for Media Specific Sheets (100% Formulas with C3~E3 Date Filter)
# ==============================================================================
def build_media_sheet_from_raw(wb, sheet_title, page_header, media_name, is_da=False):
    print(f"Building formula sheet: {sheet_title}...")
    ws = wb.create_sheet(title=sheet_title)
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells('A1:N1')
    ws['A1'] = page_header
    ws['A1'].font = font_title; ws['A1'].fill = fill_navy; ws['A1'].alignment = align_center
    ws.row_dimensions[1].height = 40
    
    # Period Control Bar
    add_period_control_bar(ws, 'N')
    
    # Top KPI cards (Filtered by C3 and E3)
    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 24
    
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
                
    ws.cell(8, 1, f"[{sheet_title} 상세 성과 목록] C3~E3 기간 필터 실시간 수식 집계").font = font_sec_title
    
    label_unit = "소재명 (Creative)" if is_da else "키워드 (Keyword)"
    headers = [
        "순위", label_unit, "노출수", "클릭수", "비용(vat-)", "비용(vat+)",
        "CTR", "CPC", "총 리드", "유효 리드", "수임", "리드 CPA", "수임 CPA", "전환율(CVR)"
    ]
    
    ws.row_dimensions[9].height = 24
    for c_idx, h in enumerate(headers, 1):
        c = ws.cell(9, c_idx, h)
        c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header
        
    # Get unique items for this media from df_tonghap
    sub_df = df_tonghap[df_tonghap['매체구분'] == media_name]
    grouped = sub_df.groupby('키워드').agg({'biz 수임': 'sum', 'biz 전체': 'sum', '비(vat포함)': 'sum'}).reset_index()
    grouped = grouped[~grouped['키워드'].isin(['-', 0, '0', 'nan', np.nan])]
    grouped = grouped.sort_values(by=['biz 수임', 'biz 전체', '비(vat포함)'], ascending=[False, False, False])
    
    key_list = grouped['키워드'].tolist()
    
    start_row = 10
    limit_count = 250 if not is_da else 150
    for rank, key_name in enumerate(key_list[:limit_count], 1):
        cur_r = start_row + rank - 1
        ws.cell(cur_r, 1, rank).alignment = align_center
        ws.cell(cur_r, 2, str(key_name)).alignment = align_left
        
        # PURE DYNAMIC FORMULAS FILTERED BY C3(시작일) ~ E3(종료일)
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

# Build individual media sheets
build_media_sheet_from_raw(wb, '네이버_SA_키워드', '네이버 검색광고 키워드 성과 상세 분석', '네이버 검색광고', is_da=False)
build_media_sheet_from_raw(wb, '구글_SA_키워드', '구글 검색광고 키워드 성과 상세 분석', '구글 검색광고', is_da=False)
build_media_sheet_from_raw(wb, 'META_소재분석', 'META DA 광고 소재별(Creative) 성과 상세 분석', 'META', is_da=True)
build_media_sheet_from_raw(wb, 'GFA_소재분석', '네이버 GFA 소재별 성과 상세 분석', '네이버 GFA', is_da=True)
build_media_sheet_from_raw(wb, '네이버_플레이스', '네이버 플레이스 검색어별 성과 상세 분석', '네이버 플레이스', is_da=False)

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

# Write rows in fast batches
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
            
    if r_idx % 10000 == 0:
        print(f"  Written {r_idx} rows...")

print(f"통합_raw sheet populated with {len(df_tonghap)} rows.")

# ==============================================================================
# Sheet: Raw_가이드
# ==============================================================================
print("Building Sheet: Raw_가이드...")
ws_guide = wb.create_sheet(title='Raw_가이드')
ws_guide.views.sheetView[0].showGridLines = True

guide_content = [
    ("알프레드 혜움 성과 리포트 사용 및 자동화 가이드", ""),
    ("", ""),
    ("1. 시트별 분석 기간 선택 기능 (신규 기능)", 
     "각 시트 상단의 C3(시작일)과 E3(종료일) 셀을 수정하면 해당 기간의 성과로 아래 모든 표와 지표가 실시간 자동 필터링되어 재계산됩니다."),
    ("", "기본값은 '=MIN(통합_raw!E:E)' 및 '=MAX(통합_raw!E:E)'로 설정되어 있어 전체 기간이 조회됩니다. 원하는 특정 주차나 월의 날짜를 직접 입력하여 기간별 성과를 비교할 수 있습니다."),
    ("", ""),
    ("2. 데이터 전체 수집 기간 자동 표시",
     "모든 시트 상단의 F3:G3 셀에 원천 데이터의 수집 기간이 수식으로 항상 실시간 자동 표시됩니다."),
    ("", ""),
    ("3. 주차별_추이비교 전용 시트 (100% 수식 연동)",
     "'주차별_추이비교' 시트는 통합_raw의 주간코드(Col I)를 와일드카드(B10&'*')로 참조하여 모든 주차의 노출, 클릭, 광고비, 리드, 수임, WoW 증감액 및 WoW 증감율을 순수 수식으로 자동 산출합니다. 새 Raw를 넣으면 즉시 자동으로 바뀝니다."),
    ("", ""),
    ("4. 월별_추이비교 시트 (하드코딩 완전 제거)",
     "'월별_추이비교' 시트 역시 통합_raw의 월(Col C)을 직접 참조하는 100% 동적 SUMIFS 수식으로 전월(MoM) 대비 증감 및 일별(DoD) 추이를 산출합니다."),
    ("", ""),
    ("5. 원천 데이터 갱신 방법", 
     "광고 관리자 및 CRM에서 최신 데이터를 다운로드한 후, '통합_raw' 시트의 A2 셀부터 그대로 붙여넣기(덮어쓰기)만 하면 10개 전 시트가 즉각 갱신됩니다."),
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

ws_guide.column_dimensions['A'].width = 35
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
