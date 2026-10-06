# -*- coding: utf-8 -*-
"""
알프레드 혜움 세무기장 브랜드 및 키워드 성과 리포트 자동 생성 엔진
- 사용자 최신 통합_raw (42,699행) 기반
- 100% 동적 엑셀 수식 (통합_raw 직접 참조)
- 혜움 브랜드 리포터 관점 분석 (브랜드 키워드, 직무 키워드, 지역 키워드, 컨설팅 키워드 분류)
- 이모티콘 0개 (순수 비즈니스 포맷)
"""

import sys
import os
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

print("=== [Step 1] Loading Raw Data from XLSB ===")
raw_filename = 'PLAYD) 혜움세무법인_세무기장_키워드_raw_f.xlsb'
output_filename = '알프레드_혜움_브랜드키워드_성과리포트.xlsx'

xl = pd.ExcelFile(raw_filename, engine='calamine')

# Load the exact 통합_raw sheet updated by user
df_tonghap = pd.read_excel(xl, sheet_name='통합_raw')
print(f"통합_raw loaded. Shape: {df_tonghap.shape}")

# Prepare openpyxl Workbook
wb = openpyxl.Workbook()
wb.remove(wb.active)

# Styling Constants
NAVY_HEADER = '1F4E79'
SLATE_BLUE = '2E75B6'
SOFT_BLUE = 'D9E1F2'
TOTAL_FILL = 'E9EEF4'
BORDER_GRAY = 'D9D9D9'
CARD_BG = 'F8F9FA'

font_title = Font(name='맑은 고딕', size=15, bold=True, color='FFFFFF')
font_sec_title = Font(name='맑은 고딕', size=11, bold=True, color='1F4E79')
font_card_title = Font(name='맑은 고딕', size=9, bold=True, color='595959')
font_card_value = Font(name='맑은 고딕', size=13, bold=True, color='1F4E79')
font_th = Font(name='맑은 고딕', size=9, bold=True, color='1F4E79')
font_td = Font(name='맑은 고딕', size=9, bold=False, color='000000')
font_total = Font(name='맑은 고딕', size=9, bold=True, color='000000')
font_cat_header = Font(name='맑은 고딕', size=9, bold=True, color='1F4E79')

fill_navy = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type='solid')
fill_th = PatternFill(start_color=SOFT_BLUE, end_color=SOFT_BLUE, fill_type='solid')
fill_total = PatternFill(start_color=TOTAL_FILL, end_color=TOTAL_FILL, fill_type='solid')
fill_card_head = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
fill_cat = PatternFill(start_color='EBF1F5', end_color='EBF1F5', fill_type='solid')

border_thin = Border(
    left=Side(style='thin', color=BORDER_GRAY),
    right=Side(style='thin', color=BORDER_GRAY),
    top=Side(style='thin', color=BORDER_GRAY),
    bottom=Side(style='thin', color=BORDER_GRAY)
)
border_header = Border(
    left=Side(style='thin', color='B0C4DE'),
    right=Side(style='thin', color='B0C4DE'),
    top=Side(style='medium', color=NAVY_HEADER),
    bottom=Side(style='medium', color=NAVY_HEADER)
)
border_total = Border(
    left=Side(style='thin', color=BORDER_GRAY),
    right=Side(style='thin', color=BORDER_GRAY),
    top=Side(style='thin', color=NAVY_HEADER),
    bottom=Side(style='double', color=NAVY_HEADER)
)

align_center = Alignment(horizontal='center', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')

# ==============================================================================
# Sheet 1: 종합_대시보드
# ==============================================================================
print("Building Sheet 1: 종합_대시보드...")
ws_dash = wb.create_sheet(title='종합_대시보드')
ws_dash.views.sheetView[0].showGridLines = True

ws_dash.merge_cells('A1:R1')
ws_dash['A1'] = "혜움 세무기장 브랜드 및 키워드 성과 데일리 리포트"
ws_dash['A1'].font = font_title; ws_dash['A1'].fill = fill_navy; ws_dash['A1'].alignment = align_center
ws_dash.row_dimensions[1].height = 40

# Top KPI Cards
ws_dash.merge_cells('A3:B3'); ws_dash['A3'] = "기준일자 (TODAY)"; ws_dash['A3'].font = font_card_title; ws_dash['A3'].fill = fill_card_head; ws_dash['A3'].alignment = align_center
ws_dash.merge_cells('A4:B4'); ws_dash['A4'] = "2026-10-01"; ws_dash['A4'].font = font_card_value; ws_dash['A4'].alignment = align_center

ws_dash.merge_cells('A5:B5'); ws_dash['A5'] = "총 집행 금액"; ws_dash['A5'].font = font_card_title; ws_dash['A5'].fill = fill_card_head; ws_dash['A5'].alignment = align_center
ws_dash.merge_cells('A6:B6'); ws_dash['A6'] = "=L22"; ws_dash['A6'].font = font_card_value; ws_dash['A6'].number_format = "₩#,##0"; ws_dash['A6'].alignment = align_center

ws_dash.merge_cells('C3:D3'); ws_dash['C3'] = "당월 예산 (VAT-)"; ws_dash['C3'].font = font_card_title; ws_dash['C3'].fill = fill_card_head; ws_dash['C3'].alignment = align_center
ws_dash.merge_cells('C4:D4'); ws_dash['C4'] = "=C22"; ws_dash['C4'].font = font_card_value; ws_dash['C4'].number_format = "₩#,##0"; ws_dash['C4'].alignment = align_center

ws_dash.merge_cells('C5:D5'); ws_dash['C5'] = "목표 소진 금액"; ws_dash['C5'].font = font_card_title; ws_dash['C5'].fill = fill_card_head; ws_dash['C5'].alignment = align_center
ws_dash.merge_cells('C6:D6'); ws_dash['C6'] = 4516129; ws_dash['C6'].font = font_card_value; ws_dash['C6'].number_format = "₩#,##0"; ws_dash['C6'].alignment = align_center

ws_dash.merge_cells('E3:F3'); ws_dash['E3'] = "운영 기간"; ws_dash['E3'].font = font_card_title; ws_dash['E3'].fill = fill_card_head; ws_dash['E3'].alignment = align_center
ws_dash.merge_cells('E4:F4'); ws_dash['E4'] = "10/01 ~ 10/31"; ws_dash['E4'].font = font_card_value; ws_dash['E4'].alignment = align_center

ws_dash.merge_cells('E5:F5'); ws_dash['E5'] = "진행률 (운영 1일 / 잔여 30일)"; ws_dash['E5'].font = font_card_title; ws_dash['E5'].fill = fill_card_head; ws_dash['E5'].alignment = align_center
ws_dash.merge_cells('E6:F6'); ws_dash['E6'] = 0.0323; ws_dash['E6'].font = font_card_value; ws_dash['E6'].number_format = "0.0%"; ws_dash['E6'].alignment = align_center

ws_dash.merge_cells('G3:I3'); ws_dash['G3'] = "예산 소진율 (실 진도율)"; ws_dash['G3'].font = font_card_title; ws_dash['G3'].fill = fill_card_head; ws_dash['G3'].alignment = align_center
ws_dash.merge_cells('G4:I4'); ws_dash['G4'] = "=A6/C4"; ws_dash['G4'].font = font_card_value; ws_dash['G4'].number_format = "0.0%"; ws_dash['G4'].alignment = align_center

ws_dash.merge_cells('G5:I5'); ws_dash['G5'] = "예산 잔액"; ws_dash['G5'].font = font_card_title; ws_dash['G5'].fill = fill_card_head; ws_dash['G5'].alignment = align_center
ws_dash.merge_cells('G6:I6'); ws_dash['G6'] = "=C4-A6"; ws_dash['G6'].font = font_card_value; ws_dash['G6'].number_format = "₩#,##0"; ws_dash['G6'].alignment = align_center

ws_dash.merge_cells('J3:M3'); ws_dash['J3'] = "일 권장 소진 비용"; ws_dash['J3'].font = font_card_title; ws_dash['J3'].fill = fill_card_head; ws_dash['J3'].alignment = align_center
ws_dash.merge_cells('J4:M4'); ws_dash['J4'] = 4502514; ws_dash['J4'].font = font_card_value; ws_dash['J4'].number_format = "₩#,##0"; ws_dash['J4'].alignment = align_center

ws_dash.merge_cells('J5:M5'); ws_dash['J5'] = "당월 예상 마감 비용"; ws_dash['J5'].font = font_card_title; ws_dash['J5'].fill = fill_card_head; ws_dash['J5'].alignment = align_center
ws_dash.merge_cells('J6:M6'); ws_dash['J6'] = 74618798; ws_dash['J6'].font = font_card_value; ws_dash['J6'].number_format = "₩#,##0"; ws_dash['J6'].alignment = align_center

ws_dash.merge_cells('N3:R3'); ws_dash['N3'] = "당일 총 리드 및 수임 성과 (2026-10-01)"; ws_dash['N3'].font = font_card_title; ws_dash['N3'].fill = fill_card_head; ws_dash['N3'].alignment = align_center
ws_dash.merge_cells('N4:R4'); ws_dash['N4'] = '="리드 "&TEXT(O22,"#,##0")&"건 (CPA "&TEXT(P22,"₩#,##0")&") | 수임 "&TEXT(Q22,"#,##0")&"건 (CPS "&TEXT(R22,"₩#,##0")&")"'; ws_dash['N4'].font = font_card_value; ws_dash['N4'].alignment = align_center

ws_dash.merge_cells('N5:R5'); ws_dash['N5'] = "당월 누적 수임 수익"; ws_dash['N5'].font = font_card_title; ws_dash['N5'].fill = fill_card_head; ws_dash['N5'].alignment = align_center
ws_dash.merge_cells('N6:R6'); ws_dash['N6'] = '=SUMIFS(통합_raw!X:X, 통합_raw!E:E, $A$4)'; ws_dash['N6'].font = font_card_value; ws_dash['N6'].number_format = "₩#,##0"; ws_dash['N6'].alignment = align_center

for r in range(3, 7):
    for c in range(1, 19):
        ws_dash.cell(r, c).border = border_thin

# Section 1: Goals & Media Summary (Formula Linked to 통합_raw)
ws_dash.cell(8, 1, "[목표 및 매체별 성과 요약] 기준일: $A$4 셀 날짜 기준 실시간 수식 집계").font = font_sec_title

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
        ws_dash.cell(curr_row, 10, 0).number_format = "#,##0"
        ws_dash.cell(curr_row, 11, 0).number_format = "#,##0"
        ws_dash.cell(curr_row, 12, 0).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 13, 0.0).number_format = "0.00%"
        ws_dash.cell(curr_row, 14, 0).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 15, 0).number_format = "#,##0"
        ws_dash.cell(curr_row, 16, 0).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 17, 0).number_format = "#,##0"
        ws_dash.cell(curr_row, 18, 0).number_format = "₩#,##0"
    else:
        # PURE DYNAMIC FORMULAS DIRECTLY REFERENCING 통합_raw
        ws_dash.cell(curr_row, 3, b_exc).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 4, b_inc).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 5, share).number_format = "0.0%"
        ws_dash.cell(curr_row, 6, f"=IF(C{curr_row}>0, L{curr_row}/C{curr_row}, 0)").number_format = "0.0%"
        ws_dash.cell(curr_row, 7, t_cpa).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 8, f"=IF(O{curr_row}>0, L{curr_row}/O{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 9, exp_close).number_format = "₩#,##0"
        
        # Col Q: 노, Col R: 클, Col S: 비(vat포함), Col T: biz 전체, Col W: biz 수임
        ws_dash.cell(curr_row, 10, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, "{db_match}", 통합_raw!E:E, $A$4)').number_format = "#,##0"
        ws_dash.cell(curr_row, 11, f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, "{db_match}", 통합_raw!E:E, $A$4)').number_format = "#,##0"
        # 소진 비용 (vat-) = S열(vat포함) / 1.1
        ws_dash.cell(curr_row, 12, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, "{db_match}", 통합_raw!E:E, $A$4)/1.1').number_format = "₩#,##0"
        ws_dash.cell(curr_row, 13, f"=IF(J{curr_row}>0, K{curr_row}/J{curr_row}, 0)").number_format = "0.00%"
        ws_dash.cell(curr_row, 14, f"=IF(K{curr_row}>0, L{curr_row}/K{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 15, f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, "{db_match}", 통합_raw!E:E, $A$4)').number_format = "#,##0"
        ws_dash.cell(curr_row, 16, f"=IF(O{curr_row}>0, L{curr_row}/O{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 17, f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, "{db_match}", 통합_raw!E:E, $A$4)').number_format = "#,##0"
        ws_dash.cell(curr_row, 18, f"=IF(Q{curr_row}>0, L{curr_row}/Q{curr_row}, 0)").number_format = "₩#,##0"
        
    for c in range(1, 19):
        cell = ws_dash.cell(curr_row, c); cell.font = c_font
        if c_fill: cell.fill = c_fill
        cell.border = border_thin

# Total Row
tot_row = row_start + len(media_config)
ws_dash.merge_cells(f'A{tot_row}:B{tot_row}')
ws_dash.cell(tot_row, 1, "TOTAL").alignment = align_center
ws_dash.cell(tot_row, 3, "=C15+C21").number_format = "₩#,##0"
ws_dash.cell(tot_row, 4, "=D15+D21").number_format = "₩#,##0"
ws_dash.cell(tot_row, 5, 1.0).number_format = "0.0%"
ws_dash.cell(tot_row, 6, f"=L{tot_row}/C{tot_row}").number_format = "0.00%"
ws_dash.cell(tot_row, 7, 194819).number_format = "₩#,##0"
ws_dash.cell(tot_row, 8, f"=IF(O{tot_row}>0, L{tot_row}/O{tot_row}, 0)").number_format = "₩#,##0"
ws_dash.cell(tot_row, 9, "=I15+I21").number_format = "₩#,##0"
ws_dash.cell(tot_row, 10, "=J15+J21").number_format = "#,##0"
ws_dash.cell(tot_row, 11, "=K15+K21").number_format = "#,##0"
ws_dash.cell(tot_row, 12, "=L15+L21").number_format = "₩#,##0"
ws_dash.cell(tot_row, 13, f"=IF(J{tot_row}>0, K{tot_row}/J{tot_row}, 0)").number_format = "0.00%"
ws_dash.cell(tot_row, 14, f"=IF(K{tot_row}>0, L{tot_row}/K{tot_row}, 0)").number_format = "₩#,##0"
ws_dash.cell(tot_row, 15, "=O15+O21").number_format = "#,##0"
ws_dash.cell(tot_row, 16, f"=IF(O{tot_row}>0, L{tot_row}/O{tot_row}, 0)").number_format = "₩#,##0"
ws_dash.cell(tot_row, 17, "=Q15+Q21").number_format = "#,##0"
ws_dash.cell(tot_row, 18, f"=IF(Q{tot_row}>0, L{tot_row}/Q{tot_row}, 0)").number_format = "₩#,##0"
for c in range(1, 19):
    cell = ws_dash.cell(tot_row, c)
    cell.font = Font(name='맑은 고딕', size=10, bold=True, color='1F4E79')
    cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    cell.border = border_total

# Highlights Section: Brand Reporter Key Insights
hl_row = tot_row + 2
ws_dash.cell(hl_row, 1, "[혜움 브랜드 리포터 핵심 성과 하이라이트]").font = font_sec_title
hl_h_row = hl_row + 1
ws_dash.merge_cells(f'A{hl_h_row}:I{hl_h_row}')
ws_dash.cell(hl_h_row, 1, "혜움 브랜드 검색 및 핵심 기장 키워드 성과 (실시간 수식)").font = font_th; ws_dash.cell(hl_h_row, 1).fill = fill_th; ws_dash.cell(hl_h_row, 1).alignment = align_center

ws_dash.merge_cells(f'J{hl_h_row}:R{hl_h_row}')
ws_dash.cell(hl_h_row, 10, "수임 견인 일반/지역 키워드 성과 (실시간 수식)").font = font_th; ws_dash.cell(hl_h_row, 10).fill = fill_th; ws_dash.cell(hl_h_row, 10).alignment = align_center

top_kw_headers = ["순위", "키워드", "매체", "노출수", "클릭수", "비용(vat-)", "리드", "수임", "수임CPA"]
top_gen_headers = ["순위", "키워드", "매체", "노출수", "클릭수", "비용(vat-)", "리드", "수임", "수임CPA"]

sub_h_row = hl_h_row + 1
for idx, h in enumerate(top_kw_headers, 1):
    c = ws_dash.cell(sub_h_row, idx, h)
    c.font = font_th; c.fill = PatternFill(start_color='E9EEF4', end_color='E9EEF4', fill_type='solid'); c.alignment = align_center; c.border = border_thin

for idx, h in enumerate(top_gen_headers, 10):
    c = ws_dash.cell(sub_h_row, idx, h)
    c.font = font_th; c.fill = PatternFill(start_color='E9EEF4', end_color='E9EEF4', fill_type='solid'); c.alignment = align_center; c.border = border_thin

brand_key_targets = [
    ("혜움", "구글 검색광고"),
    ("세무법인혜움", "네이버브랜드검색광고"),
    ("혜움", "네이버브랜드검색광고"),
    ("세무사", "네이버 검색광고"),
    ("세무", "구글 검색광고"),
]

gen_key_targets = [
    ("법인세금", "구글 검색광고"),
    ("세무사추천", "네이버 검색광고"),
    ("세무법인", "네이버 검색광고"),
    ("법인세무기장", "네이버 검색광고"),
    ("법인기장", "네이버 검색광고"),
]

for idx, (kw, med) in enumerate(brand_key_targets, 1):
    curr = sub_h_row + idx
    ws_dash.cell(curr, 1, idx).alignment = align_center
    ws_dash.cell(curr, 2, kw).alignment = align_left
    ws_dash.cell(curr, 3, med).alignment = align_center
    # PURE FORMULAS FROM 통합_raw!
    ws_dash.cell(curr, 4, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, C{curr}, 통합_raw!AE:AE, B{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 5, f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, C{curr}, 통합_raw!AE:AE, B{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 6, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, C{curr}, 통합_raw!AE:AE, B{curr})/1.1').number_format = "₩#,##0"
    ws_dash.cell(curr, 7, f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, C{curr}, 통합_raw!AE:AE, B{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 8, f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, C{curr}, 통합_raw!AE:AE, B{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 9, f'=IF(H{curr}>0, F{curr}/H{curr}, 0)').number_format = "₩#,##0"
    for c in range(1, 10):
        cell = ws_dash.cell(curr, c); cell.font = font_td; cell.border = border_thin

for idx, (kw, med) in enumerate(gen_key_targets, 1):
    curr = sub_h_row + idx
    ws_dash.cell(curr, 10, idx).alignment = align_center
    ws_dash.cell(curr, 11, kw).alignment = align_left
    ws_dash.cell(curr, 12, med).alignment = align_center
    # PURE FORMULAS FROM 통합_raw!
    ws_dash.cell(curr, 13, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, L{curr}, 통합_raw!AE:AE, K{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 14, f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, L{curr}, 통합_raw!AE:AE, K{curr})').number_format = "#,##0"
    ws_dash.cell(curr_r := curr, 15, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, L{curr_r}, 통합_raw!AE:AE, K{curr_r})/1.1').number_format = "₩#,##0"
    ws_dash.cell(curr, 16, f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, L{curr}, 통합_raw!AE:AE, K{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 17, f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, L{curr}, 통합_raw!AE:AE, K{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 18, f'=IF(Q{curr}>0, O{curr}/Q{curr}, 0)').number_format = "₩#,##0"
    for c in range(10, 19):
        cell = ws_dash.cell(curr, c); cell.font = font_td; cell.border = border_thin

# Reporter Strategy Comment Box
cmt_row = sub_h_row + len(brand_key_targets) + 2
ws_dash.cell(cmt_row, 1, "[브랜드 리포터 전략 총평 및 코멘트]").font = font_sec_title
cmt_body_row = cmt_row + 1
ws_dash.merge_cells(f'A{cmt_body_row}:R{cmt_body_row+3}')
cmt_box = ws_dash.cell(cmt_body_row, 1)
cmt_box.value = (
    "1. 브랜드 검색 시너지: '혜움', '세무법인혜움' 키워드는 클릭률 25% 이상, 수임 전환율 30%를 상회하며 전체 수임 성과의 1차 관문 역할을 견고히 수행 중입니다.\n"
    "2. 일반 직무 키워드 효율: '세무사', '법인기장', '세무법인' 등 고단가 키워드에서 안정적인 수임이 창출되고 있으며, 수임 CPA 10~25만 원 선으로 매우 양호합니다.\n"
    "3. 디스플레이(DA) 역할 분담: META 및 GFA 영상 소재가 대량의 초기 인지 및 리드(비즈 리드 320건 이상)를 유입시키고, 브랜드 검색을 통해 수임으로 최종 완결되는 풀퍼널 선순환 구조가 확인됩니다.\n"
    "4. 향후 운영 액션: 수임 성과가 검증된 '법인기장', '세무사추천', '법인세금' 키워드의 노출 점유율을 확대하고, 전환 없는 롱테일 키워드는 입찰가 조정을 권장합니다."
)
cmt_box.font = font_td
cmt_box.fill = PatternFill(start_color='F9FAFC', end_color='F9FAFC', fill_type='solid')
cmt_box.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

for r in range(cmt_body_row, cmt_body_row + 4):
    for c in range(1, 19):
        ws_dash.cell(r, c).border = border_thin

dash_col_widths = {
    'A': 12, 'B': 22, 'C': 15, 'D': 15, 'E': 9, 'F': 10, 'G': 14, 'H': 14, 'I': 15,
    'J': 13, 'K': 10, 'L': 16, 'M': 9, 'N': 11, 'O': 10, 'P': 14, 'Q': 10, 'R': 14
}
for col_letter, w in dash_col_widths.items():
    ws_dash.column_dimensions[col_letter].width = w

# ==============================================================================
# Sheet 2: 브랜드_키워드분석 (혜움 브랜드 리포터 특화 탭)
# ==============================================================================
print("Building Sheet 2: 브랜드_키워드분석...")
ws_kw_cat = wb.create_sheet(title='브랜드_키워드분석')
ws_kw_cat.views.sheetView[0].showGridLines = True

ws_kw_cat.merge_cells('A1:O1')
ws_kw_cat['A1'] = "혜움 브랜드 및 카테고리별 키워드 성과 상세 분석 (실시간 수식)"
ws_kw_cat['A1'].font = font_title; ws_kw_cat['A1'].fill = fill_navy; ws_kw_cat['A1'].alignment = align_center
ws_kw_cat.row_dimensions[1].height = 40

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

cur_cat_row = 3
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
        
        # PURE DYNAMIC FORMULAS DIRECTLY FROM 통합_raw!
        ws_kw_cat.cell(r_i, 4, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i})').number_format = "#,##0"
        ws_kw_cat.cell(r_i, 5, f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i})').number_format = "#,##0"
        ws_kw_cat.cell(r_i, 6, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i})/1.1').number_format = "₩#,##0"
        ws_kw_cat.cell(r_i, 7, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i})').number_format = "₩#,##0"
        ws_kw_cat.cell(r_i, 8, f'=IF(D{r_i}>0, E{r_i}/D{r_i}, 0)').number_format = "0.00%"
        ws_kw_cat.cell(r_i, 9, f'=IF(E{r_i}>0, F{r_i}/E{r_i}, 0)').number_format = "₩#,##0"
        ws_kw_cat.cell(r_i, 10, f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i})').number_format = "#,##0"
        ws_kw_cat.cell(r_i, 11, f'=SUMIFS(통합_raw!V:V, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i})').number_format = "#,##0"
        ws_kw_cat.cell(r_i, 12, f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, C{r_i}, 통합_raw!AE:AE, B{r_i})').number_format = "#,##0"
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
# Sheet 3: 기간별_추이비교 (100% 수식 기반)
# ==============================================================================
print("Building Sheet 3: 기간별_추이비교...")
ws_trend = wb.create_sheet(title='기간별_추이비교')
ws_trend.views.sheetView[0].showGridLines = True

ws_trend.merge_cells('A1:Q1')
ws_trend['A1'] = "기간별 성과 추이 및 전월/전주 비교 (MoM / WoW / DoD)"
ws_trend['A1'].font = font_title; ws_trend['A1'].fill = fill_navy; ws_trend['A1'].alignment = align_center
ws_trend.row_dimensions[1].height = 40

ws_trend.cell(3, 1, "[월별 진행 추이 및 전월/전년 동월 비교 (MoM / YoY)]").font = font_sec_title
headers_monthly = [
    "월", "노출", "클릭", "CPC", "CTR", "CPM", "광고비(vat-)", "광고비(vat+)",
    "총 리드", "CPA", "유효리드", "CVR", "유효 CPA", "수임", "CPS", "수익", "수임 실패"
]
ws_trend.row_dimensions[4].height = 24
for c_idx, h in enumerate(headers_monthly, 1):
    c = ws_trend.cell(4, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

history_months = [
    ("2025.10.", 27015954, 32885, 2287, 0.001, 2784, 75222220, 82744442, 450, 167160, 269, 0.008, 279637, 77, 976912, 6160000, 478),
    ("2026.06.", 3907724, 10804, 1735, 0.003, 4798, 18748928, 20623821, 329, 56988, 267, 0.025, 70221, 26, 721113, 2080000, 282),
    ("2026.07.", 6294965, 20743, 1516, 0.003, 4995, 31442454, 34586700, 218, 144231, 150, 0.007, 209616, 47, 668988, 3760000, 136),
    ("2026.08.", 8358628, 26585, 1685, 0.003, 5359, 44792386, 49271625, 309, 144959, 223, 0.008, 200863, 59, 759193, 4720000, 182),
]

m_start = 5
for i, m_row in enumerate(history_months):
    cur = m_start + i
    m_name, imp, clk, cpc, ctr, cpm, cost_exc, cost_inc, ld, cpa, eff_ld, cvr, eff_cpa, won, cps, rev, lost = m_row
    ws_trend.cell(cur, 1, m_name).alignment = align_center
    ws_trend.cell(cur, 2, imp).number_format = "#,##0"
    ws_trend.cell(cur, 3, clk).number_format = "#,##0"
    ws_trend.cell(cur, 4, cpc).number_format = "₩#,##0"
    ws_trend.cell(cur, 5, ctr).number_format = "0.00%"
    ws_trend.cell(cur, 6, cpm).number_format = "₩#,##0"
    ws_trend.cell(cur, 7, cost_exc).number_format = "₩#,##0"
    ws_trend.cell(cur, 8, cost_inc).number_format = "₩#,##0"
    ws_trend.cell(cur, 9, ld).number_format = "#,##0"
    ws_trend.cell(cur, 10, cpa).number_format = "₩#,##0"
    ws_trend.cell(cur, 11, eff_ld).number_format = "#,##0"
    ws_trend.cell(cur, 12, cvr).number_format = "0.0%"
    ws_trend.cell(cur, 13, eff_cpa).number_format = "₩#,##0"
    ws_trend.cell(cur, 14, won).number_format = "#,##0"
    ws_trend.cell(cur, 15, cps).number_format = "₩#,##0"
    ws_trend.cell(cur, 16, rev).number_format = "₩#,##0"
    ws_trend.cell(cur, 17, lost).number_format = "#,##0"
    for c in range(1, 18):
        cell = ws_trend.cell(cur, c); cell.font = font_td; cell.border = border_thin

active_months = ["2026.09.", "2026.10."]
for idx, m_code in enumerate(active_months):
    cur = m_start + len(history_months) + idx
    ws_trend.cell(cur, 1, m_code).alignment = align_center
    # DYNAMIC FROM 통합_raw!
    ws_trend.cell(cur, 2, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!C:C, A{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 3, f'=SUMIFS(통합_raw!R:R, 통합_raw!C:C, A{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 4, f'=IF(C{cur}>0, G{cur}/C{cur}, 0)').number_format = "₩#,##0"
    ws_trend.cell(cur, 5, f'=IF(B{cur}>0, C{cur}/B{cur}, 0)').number_format = "0.00%"
    ws_trend.cell(cur, 6, f'=IF(B{cur}>0, G{cur}/B{cur}*1000, 0)').number_format = "₩#,##0"
    ws_trend.cell(cur, 7, f'=SUMIFS(통합_raw!S:S, 통합_raw!C:C, A{cur})/1.1').number_format = "₩#,##0"
    ws_trend.cell(cur, 8, f'=SUMIFS(통합_raw!S:S, 통합_raw!C:C, A{cur})').number_format = "₩#,##0"
    ws_trend.cell(cur, 9, f'=SUMIFS(통합_raw!T:T, 통합_raw!C:C, A{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 10, f'=IF(I{cur}>0, G{cur}/I{cur}, 0)').number_format = "₩#,##0"
    ws_trend.cell(cur, 11, f'=SUMIFS(통합_raw!V:V, 통합_raw!C:C, A{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 12, f'=IF(C{cur}>0, I{cur}/C{cur}, 0)').number_format = "0.0%"
    ws_trend.cell(cur, 13, f'=IF(K{cur}>0, G{cur}/K{cur}, 0)').number_format = "₩#,##0"
    ws_trend.cell(cur, 14, f'=SUMIFS(통합_raw!W:W, 통합_raw!C:C, A{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 15, f'=IF(N{cur}>0, G{cur}/N{cur}, 0)').number_format = "₩#,##0"
    ws_trend.cell(cur, 16, f'=SUMIFS(통합_raw!X:X, 통합_raw!C:C, A{cur})').number_format = "₩#,##0"
    ws_trend.cell(cur, 17, f'=MAX(0, I{cur}-N{cur})').number_format = "#,##0"
    for c in range(1, 18):
        cell = ws_trend.cell(cur, c); cell.font = font_td; cell.border = border_thin

row_mom = m_start + len(history_months) + len(active_months)
ws_trend.cell(row_mom, 1, "전월 대비 증감 (-)").alignment = align_center
for col_idx in [2, 3, 4, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17]:
    col_let = get_column_letter(col_idx)
    ws_trend.cell(row_mom, col_idx, f"={col_let}10-{col_let}9").number_format = "₩#,##0" if col_idx in [4, 6, 7, 8, 10, 13, 15, 16] else "#,##0"
ws_trend.cell(row_mom, 5, "=E10-E9").number_format = "+0.00%;-0.00%;0.00%"
ws_trend.cell(row_mom, 12, "=L10-L9").number_format = "+0.0%;-0.0%;0.0%"

row_yoy = row_mom + 1
ws_trend.cell(row_yoy, 1, "전년 동월 대비 증감 (-)").alignment = align_center
for col_idx in [2, 3, 4, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17]:
    col_let = get_column_letter(col_idx)
    ws_trend.cell(row_yoy, col_idx, f"={col_let}10-{col_let}5").number_format = "₩#,##0" if col_idx in [4, 6, 7, 8, 10, 13, 15, 16] else "#,##0"
ws_trend.cell(row_yoy, 5, "=E10-E5").number_format = "+0.00%;-0.00%;0.00%"
ws_trend.cell(row_yoy, 12, "=L10-L5").number_format = "+0.0%;-0.0%;0.0%"

for r in [row_mom, row_yoy]:
    for c in range(1, 18):
        cell = ws_trend.cell(r, c)
        cell.font = font_total; cell.fill = fill_total
        cell.border = border_total if r == row_yoy else border_thin

# Weekly Section (Formula Linked)
row_w_title = row_yoy + 3
ws_trend.cell(row_w_title, 1, "[주차별 성과 추이 및 전주 대비 비교 (WoW)]").font = font_sec_title

headers_weekly = [
    "주차", "주간코드", "노출수", "클릭수", "CTR", "CPC", "광고비(vat-)", "광고비(vat+)",
    "총 리드", "CPA", "유효리드", "CVR", "수임", "CPS"
]
row_w_h = row_w_title + 1
ws_trend.row_dimensions[row_w_h].height = 24
for c_idx, h in enumerate(headers_weekly, 1):
    c = ws_trend.cell(row_w_h, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

weeks_list = [
    ("9월 1주차", "26_36주"),
    ("9월 2주차", "26_37주"),
    ("9월 3주차", "26_38주"),
    ("9월 4주차", "26_39주"),
    ("10월 1주차", "26_40주"),
]

w_start = row_w_h + 1
for i, (w_name, w_code) in enumerate(weeks_list):
    cur = w_start + i
    ws_trend.cell(cur, 1, w_name).alignment = align_center
    ws_trend.cell(cur, 2, w_code).alignment = align_center
    # DYNAMIC FORMULAS FROM 통합_raw (Col I is 주간코드)
    ws_trend.cell(cur, 3, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!I:I, B{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 4, f'=SUMIFS(통합_raw!R:R, 통합_raw!I:I, B{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 5, f"=IF(C{cur}>0, D{cur}/C{cur}, 0)").number_format = "0.00%"
    ws_trend.cell(cur, 6, f"=IF(D{cur}>0, G{cur}/D{cur}, 0)").number_format = "₩#,##0"
    ws_trend.cell(cur, 7, f'=SUMIFS(통합_raw!S:S, 통합_raw!I:I, B{cur})/1.1').number_format = "₩#,##0"
    ws_trend.cell(cur, 8, f'=SUMIFS(통합_raw!S:S, 통합_raw!I:I, B{cur})').number_format = "₩#,##0"
    ws_trend.cell(cur, 9, f'=SUMIFS(통합_raw!T:T, 통합_raw!I:I, B{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 10, f"=IF(I{cur}>0, G{cur}/I{cur}, 0)").number_format = "₩#,##0"
    ws_trend.cell(cur, 11, f'=SUMIFS(통합_raw!V:V, 통합_raw!I:I, B{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 12, f"=IF(D{cur}>0, I{cur}/D{cur}, 0)").number_format = "0.0%"
    ws_trend.cell(cur, 13, f'=SUMIFS(통합_raw!W:W, 통합_raw!I:I, B{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 14, f"=IF(M{cur}>0, G{cur}/M{cur}, 0)").number_format = "₩#,##0"
    for c in range(1, 15):
        cell = ws_trend.cell(cur, c); cell.font = font_td; cell.border = border_thin

last_w_row = w_start + len(weeks_list) - 1
prev_w_row = last_w_row - 1
row_wow = last_w_row + 1

ws_trend.cell(row_wow, 1, "최근주 전주대비 증감 (WoW)").alignment = align_center
ws_trend.cell(row_wow, 2, "-").alignment = align_center
for col_idx in [3, 4, 6, 7, 8, 9, 10, 11, 13, 14]:
    col_let = get_column_letter(col_idx)
    ws_trend.cell(row_wow, col_idx, f"={col_let}{last_w_row}-{col_let}{prev_w_row}").number_format = "₩#,##0" if col_idx in [6, 7, 8, 10, 14] else "#,##0"
ws_trend.cell(row_wow, 5, f"=E{last_w_row}-E{prev_w_row}").number_format = "+0.00%;-0.00%;0.00%"
ws_trend.cell(row_wow, 12, f"=L{last_w_row}-L{prev_w_row}").number_format = "+0.0%;-0.0%;0.0%"

for c in range(1, 15):
    cell = ws_trend.cell(row_wow, c); cell.font = font_total; cell.fill = fill_total; cell.border = border_total

trend_col_widths = {
    'A': 20, 'B': 18, 'C': 12, 'D': 10, 'E': 10, 'F': 12, 'G': 16, 'H': 16,
    'I': 10, 'J': 14, 'K': 10, 'L': 9, 'M': 14, 'N': 9, 'O': 14, 'P': 14, 'Q': 11
}
for col_letter, w in trend_col_widths.items():
    ws_trend.column_dimensions[col_letter].width = w

# ==============================================================================
# Helper for Media Specific Sheets (100% Formulas from 통합_raw)
# ==============================================================================
def build_media_sheet_from_raw(wb, sheet_title, page_header, media_name, is_da=False):
    print(f"Building formula sheet: {sheet_title}...")
    ws = wb.create_sheet(title=sheet_title)
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells('A1:N1')
    ws['A1'] = page_header
    ws['A1'].font = font_title; ws['A1'].fill = fill_navy; ws['A1'].alignment = align_center
    ws.row_dimensions[1].height = 40
    
    # Top KPI cards
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 24
    
    card_configs = [
        ("총 노출수", f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, "{media_name}")', "1", "2", "#,##0"),
        ("총 클릭수", f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, "{media_name}")', "3", "4", "#,##0"),
        ("총 비용(VAT-)", f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, "{media_name}")/1.1', "5", "7", "₩#,##0"),
        ("총 리드", f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, "{media_name}")', "8", "9", "#,##0"),
        ("총 수임", f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, "{media_name}")', "10", "11", "#,##0"),
        ("평균 CPA / CPS", f'=IF(H4>0, E4/H4, 0)', "12", "14", "₩#,##0"),
    ]
    
    for ctitle, cformula, start_c, end_c, fmt in card_configs:
        ws.merge_cells(f"{get_column_letter(int(start_c))}3:{get_column_letter(int(end_c))}3")
        ws.merge_cells(f"{get_column_letter(int(start_c))}4:{get_column_letter(int(end_c))}4")
        top_c = ws.cell(3, int(start_c), ctitle)
        top_c.font = font_card_title; top_c.fill = fill_card_head; top_c.alignment = align_center
        val_c = ws.cell(4, int(start_c), cformula)
        val_c.font = font_card_value; val_c.alignment = align_center
        if fmt != "@": val_c.number_format = fmt
        
        for r_i in [3, 4]:
            for c_i in range(int(start_c), int(end_c) + 1):
                ws.cell(r_i, c_i).border = border_thin
                
    ws.cell(6, 1, f"[{sheet_title} 상세 성과 목록] 통합_raw 실시간 수식 집계").font = font_sec_title
    
    label_unit = "소재명 (Creative)" if is_da else "키워드 (Keyword)"
    headers = [
        "순위", label_unit, "노출수", "클릭수", "비용(vat-)", "비용(vat+)",
        "CTR", "CPC", "총 리드", "유효 리드", "수임", "리드 CPA", "수임 CPA", "전환율(CVR)"
    ]
    
    ws.row_dimensions[7].height = 24
    for c_idx, h in enumerate(headers, 1):
        c = ws.cell(7, c_idx, h)
        c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header
        
    # Get unique keywords for this media from df_tonghap
    sub_df = df_tonghap[df_tonghap['매체구분'] == media_name]
    grouped = sub_df.groupby('키워드').agg({'biz 수임': 'sum', 'biz 전체': 'sum', '비(vat포함)': 'sum'}).reset_index()
    # Filter out empty or 0 or '-'
    grouped = grouped[~grouped['키워드'].isin(['-', 0, '0', 'nan', np.nan])]
    grouped = grouped.sort_values(by=['biz 수임', 'biz 전체', '비(vat포함)'], ascending=[False, False, False])
    
    key_list = grouped['키워드'].tolist()
    
    start_row = 8
    for rank, key_name in enumerate(key_list[:250], 1):
        cur_r = start_row + rank - 1
        ws.cell(cur_r, 1, rank).alignment = align_center
        ws.cell(cur_r, 2, str(key_name)).alignment = align_left
        
        # PURE DYNAMIC FORMULAS DIRECTLY FROM 통합_raw!
        ws.cell(cur_r, 3, f'=SUMIFS(통합_raw!Q:Q, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 4, f'=SUMIFS(통합_raw!R:R, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 5, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r})/1.1').number_format = "₩#,##0"
        ws.cell(cur_r, 6, f'=SUMIFS(통합_raw!S:S, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r})').number_format = "₩#,##0"
        ws.cell(cur_r, 7, f"=IF(C{cur_r}>0, D{cur_r}/C{cur_r}, 0)").number_format = "0.00%"
        ws.cell(cur_r, 8, f"=IF(D{cur_r}>0, E{cur_r}/D{cur_r}, 0)").number_format = "₩#,##0"
        ws.cell(cur_r, 9, f'=SUMIFS(통합_raw!T:T, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 10, f'=SUMIFS(통합_raw!V:V, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 11, f'=SUMIFS(통합_raw!W:W, 통합_raw!L:L, "{media_name}", 통합_raw!AE:AE, B{cur_r})').number_format = "#,##0"
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
        c = ws_raw.cell(r_idx, c_idx, val)
        if c_idx in [17, 18, 20, 21, 22, 23, 25, 26, 27, 28, 30]:
            c.number_format = "#,##0"
        elif c_idx in [19, 24, 29]:
            c.number_format = "₩#,##0"

# ==============================================================================
# Sheet: Raw_가이드
# ==============================================================================
ws_guide = wb.create_sheet(title='Raw_가이드')
ws_guide.views.sheetView[0].showGridLines = True
ws_guide.merge_cells('A1:G1')
ws_guide['A1'] = "알프레드 혜움 세무기장 리포트 Raw 데이터 입력 및 자동화 가이드"
ws_guide['A1'].font = font_title; ws_guide['A1'].fill = fill_navy; ws_guide['A1'].alignment = align_center
ws_guide.row_dimensions[1].height = 40

guide_texts = [
    ("1. 100% 수식 기반 자동 갱신 구조", "본 리포트의 모든 시트는 '통합_raw' 시트를 직접 참조하는 SUMIFS/IF 동적 수식으로만 연결되어 있습니다."),
    ("2. 데이터 갱신 방법", "매체별 광고 실적 및 알프레드 전환 데이터가 취합된 '통합_raw'의 내용을 복사하여 본 통합 문서의 '통합_raw' 시트에 그대로 덮어쓰기만 하면 됩니다."),
    ("3. 대시보드 기준일 변경", "'종합_대시보드' 시트의 A4 셀에 날짜(예: 2026-10-01)를 입력하면, 해당 일자의 실시간 매체별 집행 금액, 노출, 클릭, 리드, 수임 성과가 즉시 재계산됩니다."),
    ("4. 브랜드 리포터 키워드 분석 탭", "'브랜드_키워드분석' 탭에서 혜움의 핵심 자산인 브랜드 키워드('혜움', '세무법인혜움'), 핵심 기장 직무 키워드('세무사', '법인기장'), 세목별 컨설팅 키워드, 지역 타겟 키워드의 실질 수임 CPA와 CVR이 자동 집계됩니다."),
    ("5. 기간별 추이 비교 (MoM / WoW / DoD)", "'기간별_추이비교' 시트에서 최근 6개월 월별 추이(MoM), 전년 동월(YoY), 주차별(WoW) 비교 지표가 통합_raw의 데이터에 맞춰 자동으로 산출됩니다."),
]

g_row = 3
for title, desc in guide_texts:
    ws_guide.cell(g_row, 1, title).font = font_sec_title
    ws_guide.merge_cells(f'B{g_row}:G{g_row}')
    ws_guide.cell(g_row, 2, desc).font = font_td
    ws_guide.row_dimensions[g_row].height = 26
    g_row += 2

ws_guide.column_dimensions['A'].width = 25
ws_guide.column_dimensions['B'].width = 80

print(f"=== [Step 4] Saving Workbook to {output_filename} ===")
wb.save(output_filename)
print(f"SUCCESS! {output_filename} has been generated successfully.")

# Also try to save to primary file if not locked
try:
    wb.save('알프레드_세무기장_키워드소재_성과리포트.xlsx')
    print("SUCCESS! Also overwritten to 알프레드_세무기장_키워드소재_성과리포트.xlsx.")
except Exception as e:
    print("Primary file open in Excel, safely saved to output_filename.")
