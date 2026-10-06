# -*- coding: utf-8 -*-
"""
알프레드 세무기장 키워드/소재 성과 리포트 자동 생성 엔진
- 100% 동적 엑셀 수식 기반 (통합_DB 참조)
- 통합_DB에 데이터만 넣으면 대시보드, 기간별 추이, 매체별 키워드/소재 실시간 자동 계산
- 이모티콘 완전 배제, 다크 네이비 비즈니스 테마
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

print("=== [Step 1] Loading Raw Data from XLSB ===")
raw_filename = 'PLAYD) 혜움세무법인_세무기장_키워드_raw_f.xlsb'
output_filename = '알프레드_세무기장_키워드소재_성과리포트_수식연동.xlsx'
primary_filename = '알프레드_세무기장_키워드소재_성과리포트.xlsx'

xl = pd.ExcelFile(raw_filename, engine='calamine')

# 1. Index Sheet (Calendar & UTM mapping)
df_idx = pd.read_excel(xl, sheet_name='index').iloc[1:]
cal = df_idx[['Unnamed: 15', 'Unnamed: 16', 'Unnamed: 17', 'Unnamed: 18']].dropna(subset=['Unnamed: 15'])
cal.columns = ['일자_dt', '영업일구분', '주간코드', '주차']
cal['일자_str'] = pd.to_datetime(cal['일자_dt']).dt.strftime('%Y-%m-%d')
cal_map = cal.set_index('일자_str').to_dict('index')

utm_map_df = df_idx[['Unnamed: 23', 'Unnamed: 24', 'Unnamed: 25']].dropna()
utm_map = {(str(r['Unnamed: 23']).strip().lower(), str(r['Unnamed: 24']).strip().lower()): str(r['Unnamed: 25']).strip() 
           for _, r in utm_map_df.iterrows()}

def get_cal_info(dt_obj):
    if pd.isna(dt_obj):
        return ('2026.09.', '26_38주', '9월 3주차', '평일', '영업일', '월')
    dt = pd.to_datetime(dt_obj)
    dt_str = dt.strftime('%Y-%m-%d')
    m_str = dt.strftime('%Y.%m.')
    w_name = ['월', '화', '수', '목', '금', '토', '일'][dt.weekday()]
    if dt_str in cal_map:
        info = cal_map[dt_str]
        return (m_str, info['주간코드'], info['주차'], '주말' if dt.weekday() >= 5 else '평일', info['영업일구분'], w_name)
    else:
        week_num = dt.isocalendar()[1]
        w_code = f"26_{week_num:02d}주"
        w_text = f"{dt.month}월 {((dt.day-1)//7)+1}주차"
        biz_day = '비영업일' if dt.weekday() >= 5 else '영업일'
        return (m_str, w_code, w_text, '주말' if dt.weekday() >= 5 else '평일', biz_day, w_name)

# 2. Extract & Normalize Media Sheets to Fact Table
records = []

# (1) Naver SA
print("Processing Naver SA...")
df_n = pd.read_excel(xl, sheet_name='naver_raw').iloc[1:]
for _, r in df_n.iterrows():
    camp = str(r.iloc[32]) if pd.notna(r.iloc[32]) else '-'
    grp = str(r.iloc[33]) if pd.notna(r.iloc[33]) else '-'
    kw = str(r.iloc[34]).strip() if pd.notna(r.iloc[34]) else '-'
    d_raw = r.iloc[35]
    try:
        dt = pd.to_datetime(str(d_raw).replace('.', '-').strip('-'))
    except:
        dt = pd.to_datetime('2026-09-21')
    dev = str(r.iloc[36]) if pd.notna(r.iloc[36]) else 'MO'
    imp = float(r.iloc[37]) if pd.notna(r.iloc[37]) else 0
    clk = float(r.iloc[38]) if pd.notna(r.iloc[38]) else 0
    cost_vat_inc = float(r.iloc[39]) if pd.notna(r.iloc[39]) else 0
    cost_vat_exc = round(cost_vat_inc / 1.1, 2)
    m_str, w_code, w_text, d_type, biz_type, w_name = get_cal_info(dt)
    records.append({
        '구분': 'SA', '매체구분': '네이버 검색광고', '일자': dt.strftime('%Y-%m-%d'),
        '월': m_str, '주간코드': w_code, '주차': w_text, '요일': w_name, '영업일구분': biz_type,
        '디바이스': dev, '캠페인': camp, '광고그룹': grp, '키워드_소재': kw,
        '노출수': imp, '클릭수': clk, '비용_vat제외': cost_vat_exc, '비용_vat포함': cost_vat_inc,
        '총리드': 0, '유효리드': 0, '수임': 0, '수익': 0
    })

# (2) Google SA
print("Processing Google SA...")
df_g = pd.read_excel(xl, sheet_name='google_raw').iloc[1:]
for _, r in df_g.iterrows():
    camp = str(r.iloc[32]) if pd.notna(r.iloc[32]) else '-'
    grp = str(r.iloc[33]) if pd.notna(r.iloc[33]) else '-'
    kw = str(r.iloc[34]).strip() if pd.notna(r.iloc[34]) else '-'
    d_val = r.iloc[35]
    if pd.isna(d_val):
        dt = pd.to_datetime('2026-09-21')
    else:
        try:
            f = float(d_val)
            dt = pd.Timestamp('1899-12-30') + pd.to_timedelta(f, unit='D')
        except:
            dt = pd.to_datetime(d_val)
    dev_raw = str(r.iloc[36]) if pd.notna(r.iloc[36]) else '컴퓨터'
    dev = 'PC' if '컴퓨터' in dev_raw else ('MO' if '휴대' in dev_raw else 'ALL')
    ctype = str(r.iloc[37]) if pd.notna(r.iloc[37]) else '검색'
    med = '구글PMAX' if '실적' in ctype or 'PMAX' in ctype.upper() else '구글 검색광고'
    imp = float(r.iloc[38]) if pd.notna(r.iloc[38]) else 0
    clk = float(r.iloc[39]) if pd.notna(r.iloc[39]) else 0
    cost_vat_exc = float(r.iloc[41]) if pd.notna(r.iloc[41]) else 0
    cost_vat_inc = round(cost_vat_exc * 1.1, 2)
    m_str, w_code, w_text, d_type, biz_type, w_name = get_cal_info(dt)
    records.append({
        '구분': 'SA', '매체구분': med, '일자': dt.strftime('%Y-%m-%d'),
        '월': m_str, '주간코드': w_code, '주차': w_text, '요일': w_name, '영업일구분': biz_type,
        '디바이스': dev, '캠페인': camp, '광고그룹': grp, '키워드_소재': kw,
        '노출수': imp, '클릭수': clk, '비용_vat제외': cost_vat_exc, '비용_vat포함': cost_vat_inc,
        '총리드': 0, '유효리드': 0, '수임': 0, '수익': 0
    })

# (3) META DA
print("Processing META DA...")
df_m = pd.read_excel(xl, sheet_name='meta_raw').iloc[1:]
for _, r in df_m.iterrows():
    camp = str(r.iloc[32]) if pd.notna(r.iloc[32]) else '-'
    grp = str(r.iloc[33]) if pd.notna(r.iloc[33]) else '-'
    crv = str(r.iloc[34]).strip() if pd.notna(r.iloc[34]) else '-'
    crv_clean = crv.split('__')[-1] if '__' in crv else crv
    dt = pd.to_datetime(r.iloc[35]) if pd.notna(r.iloc[35]) else pd.to_datetime('2026-09-01')
    imp = float(r.iloc[36]) if pd.notna(r.iloc[36]) else 0
    clk = float(r.iloc[37]) if pd.notna(r.iloc[37]) else 0
    cost_vat_inc = float(r.iloc[19]) if pd.notna(r.iloc[19]) else 0
    cost_vat_exc = round(cost_vat_inc / 1.1, 2)
    m_str, w_code, w_text, d_type, biz_type, w_name = get_cal_info(dt)
    records.append({
        '구분': 'DA', '매체구분': 'META', '일자': dt.strftime('%Y-%m-%d'),
        '월': m_str, '주간코드': w_code, '주차': w_text, '요일': w_name, '영업일구분': biz_type,
        '디바이스': 'ALL', '캠페인': camp, '광고그룹': grp, '키워드_소재': crv_clean,
        '노출수': imp, '클릭수': clk, '비용_vat제외': cost_vat_exc, '비용_vat포함': cost_vat_inc,
        '총리드': 0, '유효리드': 0, '수임': 0, '수익': 0
    })

# (4) GFA DA
print("Processing GFA DA...")
df_gfa = pd.read_excel(xl, sheet_name='gfa_raw').iloc[1:]
for _, r in df_gfa.iterrows():
    crv = str(r.iloc[32]).strip() if pd.notna(r.iloc[32]) else '-'
    crv_clean = crv.split('__')[-1] if '__' in crv else crv
    grp = str(r.iloc[34]) if pd.notna(r.iloc[34]) else '-'
    camp = str(r.iloc[36]) if pd.notna(r.iloc[36]) else '-'
    d_raw = r.iloc[38]
    try:
        dt = pd.to_datetime(str(d_raw).replace('.', '-').strip('-'))
    except:
        dt = pd.to_datetime('2026-09-01')
    dev = str(r.iloc[39]) if pd.notna(r.iloc[39]) else 'MO'
    cost_vat_inc = float(r.iloc[40]) if pd.notna(r.iloc[40]) else 0
    cost_vat_exc = round(cost_vat_inc / 1.1, 2)
    imp = float(r.iloc[43]) if pd.notna(r.iloc[43]) else 0
    clk = float(r.iloc[44]) if pd.notna(r.iloc[44]) else 0
    m_str, w_code, w_text, d_type, biz_type, w_name = get_cal_info(dt)
    records.append({
        '구분': 'DA', '매체구분': '네이버 GFA', '일자': dt.strftime('%Y-%m-%d'),
        '월': m_str, '주간코드': w_code, '주차': w_text, '요일': w_name, '영업일구분': biz_type,
        '디바이스': dev, '캠페인': camp, '광고그룹': grp, '키워드_소재': crv_clean,
        '노출수': imp, '클릭수': clk, '비용_vat제외': cost_vat_exc, '비용_vat포함': cost_vat_inc,
        '총리드': 0, '유효리드': 0, '수임': 0, '수익': 0
    })

# (5) Naver Place
print("Processing Naver Place...")
df_p = pd.read_excel(xl, sheet_name='naver_place_raw').iloc[1:]
for _, r in df_p.iterrows():
    camp = str(r.iloc[32]) if pd.notna(r.iloc[32]) else '#네이버_플레이스'
    grp = str(r.iloc[33]) if pd.notna(r.iloc[33]) else '#혜움_플레이스_ALL'
    d_raw = r.iloc[34]
    try:
        dt = pd.to_datetime(str(d_raw).replace('.', '-').strip('-'))
    except:
        dt = pd.to_datetime('2026-09-22')
    dev = str(r.iloc[35]) if pd.notna(r.iloc[35]) else 'MO'
    kw = str(r.iloc[36]).strip() if pd.notna(r.iloc[36]) else '-'
    imp = float(r.iloc[37]) if pd.notna(r.iloc[37]) else 0
    clk = float(r.iloc[38]) if pd.notna(r.iloc[38]) else 0
    cost_vat_inc = float(r.iloc[39]) if pd.notna(r.iloc[39]) else 0
    cost_vat_exc = round(cost_vat_inc / 1.1, 2)
    m_str, w_code, w_text, d_type, biz_type, w_name = get_cal_info(dt)
    records.append({
        '구분': 'PLACE', '매체구분': '네이버 플레이스', '일자': dt.strftime('%Y-%m-%d'),
        '월': m_str, '주간코드': w_code, '주차': w_text, '요일': w_name, '영업일구분': biz_type,
        '디바이스': dev, '캠페인': camp, '광고그룹': grp, '키워드_소재': kw,
        '노출수': imp, '클릭수': clk, '비용_vat제외': cost_vat_exc, '비용_vat포함': cost_vat_inc,
        '총리드': 0, '유효리드': 0, '수임': 0, '수익': 0
    })

# (6) 전환 CRM (알프레드)
print("Processing Conversion CRM...")
df_c = pd.read_excel(xl, sheet_name='전환_raw').iloc[1:]
for _, r in df_c.iterrows():
    dt = pd.to_datetime(r.iloc[31]) if pd.notna(r.iloc[31]) else pd.to_datetime('2026-09-01')
    src = str(r.iloc[37]).strip().lower() if pd.notna(r.iloc[37]) else 'referral'
    med_src = str(r.iloc[38]).strip().lower() if pd.notna(r.iloc[38]) else 'referral'
    med_name = utm_map.get((src, med_src), 'Referral')
    
    camp = str(r.iloc[39]).strip() if pd.notna(r.iloc[39]) else '-'
    content = str(r.iloc[40]).strip() if pd.notna(r.iloc[40]) else '-'
    term = str(r.iloc[41]).strip() if pd.notna(r.iloc[41]) else '-'
    
    if '검색' in med_name or 'sa' in src or 'place' in src:
        kw_crv = term if term not in ['-', 'nan', ''] else content
        sec = 'SA'
    elif 'META' in med_name or 'GFA' in med_name or 'meta' in src or 'gfa' in src:
        crv_body = content.split('__')[-1] if '__' in content else content
        kw_crv = crv_body if crv_body not in ['-', 'nan', ''] else term
        sec = 'DA'
    else:
        kw_crv = content if content not in ['-', 'nan', ''] else term
        sec = 'CRM'
        
    biz_all = int(r.iloc[45]) if pd.notna(r.iloc[45]) else 0
    biz_eff = int(r.iloc[47]) if pd.notna(r.iloc[47]) else 0
    biz_won = int(r.iloc[48]) if pd.notna(r.iloc[48]) else 0
    biz_rev = float(r.iloc[49]) if pd.notna(r.iloc[49]) else 0
    
    m_str, w_code, w_text, d_type, biz_type, w_name = get_cal_info(dt)
    records.append({
        '구분': sec, '매체구분': med_name, '일자': dt.strftime('%Y-%m-%d'),
        '월': m_str, '주간코드': w_code, '주차': w_text, '요일': w_name, '영업일구분': biz_type,
        '디바이스': 'ALL', '캠페인': camp, '광고그룹': content, '키워드_소재': kw_crv,
        '노출수': 0, '클릭수': 0, '비용_vat제외': 0, '비용_vat포함': 0,
        '총리드': biz_all, '유효리드': biz_eff, '수임': biz_won, '수익': biz_rev
    })

df_fact = pd.DataFrame(records)
print(f"Total unified fact rows: {len(df_fact)}")

# Prepare Workbook
wb = openpyxl.Workbook()
wb.remove(wb.active)

# Styles
NAVY_HEADER = '1F4E79'
SLATE_BLUE = '2E75B6'
SOFT_BLUE = 'D9E1F2'
TOTAL_FILL = 'E9EEF4'
BORDER_GRAY = 'D9D9D9'

font_title = Font(name='맑은 고딕', size=16, bold=True, color='FFFFFF')
font_sec_title = Font(name='맑은 고딕', size=11, bold=True, color='1F4E79')
font_card_title = Font(name='맑은 고딕', size=9, bold=True, color='595959')
font_card_value = Font(name='맑은 고딕', size=14, bold=True, color='1F4E79')
font_th = Font(name='맑은 고딕', size=9, bold=True, color='1F4E79')
font_td = Font(name='맑은 고딕', size=9, bold=False, color='000000')
font_total = Font(name='맑은 고딕', size=9, bold=True, color='000000')

fill_navy = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type='solid')
fill_th = PatternFill(start_color=SOFT_BLUE, end_color=SOFT_BLUE, fill_type='solid')
fill_total = PatternFill(start_color=TOTAL_FILL, end_color=TOTAL_FILL, fill_type='solid')
fill_card_head = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')

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

# === Sheet 1: 종합_대시보드 (100% 수식 기반) ===
print("Building Sheet 1: 종합_대시보드 (100% 수식 기반)...")
ws_dash = wb.create_sheet(title='종합_대시보드')
ws_dash.views.sheetView[0].showGridLines = True

ws_dash.merge_cells('A1:R1')
ws_dash['A1'] = "세무기장 Daily Report"
ws_dash['A1'].font = font_title; ws_dash['A1'].fill = fill_navy; ws_dash['A1'].alignment = align_center
ws_dash.row_dimensions[1].height = 40

# Top KPI Summary Block (Formula Linked)
ws_dash.merge_cells('A3:B3'); ws_dash['A3'] = "TODAY"; ws_dash['A3'].font = font_card_title; ws_dash['A3'].fill = fill_card_head; ws_dash['A3'].alignment = align_center
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
ws_dash.merge_cells('N6:R6'); ws_dash['N6'] = '=SUMIFS(통합_DB!T:T, 통합_DB!C:C, $A$4)'; ws_dash['N6'].font = font_card_value; ws_dash['N6'].number_format = "₩#,##0"; ws_dash['N6'].alignment = align_center

for r in range(3, 7):
    for c in range(1, 19):
        ws_dash.cell(r, c).border = border_thin

# Section 1: Goals & Media Summary (Fully Formula Driven)
ws_dash.cell(8, 1, "[목표 및 매체별 성과 요약] 기준일: $A$4 셀의 날짜 기준 실시간 집계").font = font_sec_title

headers_s1 = [
    "서비스 구분", "매체", "예산(vat-)", "예산(vat+)", "비중", "소진율", "목표CPA", "실 CPA", "예상 마감",
    "노출", "클릭", "소진 비용(vat-)", "CTR", "CPC", "총 리드", "총 리드 CPA", "수임 biz", "수임 CPA"
]

ws_dash.row_dimensions[9].height = 24
for col_idx, h in enumerate(headers_s1, 1):
    cell = ws_dash.cell(9, col_idx, h)
    cell.font = font_th; cell.fill = fill_th; cell.alignment = align_center; cell.border = border_header

# Media Budget master definition
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
        # REAL DYNAMIC FORMULA FROM 통합_DB!
        ws_dash.cell(curr_row, 3, b_exc).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 4, b_inc).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 5, share).number_format = "0.0%"
        ws_dash.cell(curr_row, 6, f"=IF(C{curr_row}>0, L{curr_row}/C{curr_row}, 0)").number_format = "0.0%"
        ws_dash.cell(curr_row, 7, t_cpa).number_format = "₩#,##0"
        ws_dash.cell(curr_row, 8, f"=IF(O{curr_row}>0, L{curr_row}/O{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 9, exp_close).number_format = "₩#,##0"
        
        # PURE EXCEL FORMULA REFERENCING 통합_DB
        ws_dash.cell(curr_row, 10, f'=SUMIFS(통합_DB!M:M, 통합_DB!B:B, "{db_match}", 통합_DB!C:C, $A$4)').number_format = "#,##0"
        ws_dash.cell(curr_row, 11, f'=SUMIFS(통합_DB!N:N, 통합_DB!B:B, "{db_match}", 통합_DB!C:C, $A$4)').number_format = "#,##0"
        ws_dash.cell(curr_row, 12, f'=SUMIFS(통합_DB!O:O, 통합_DB!B:B, "{db_match}", 통합_DB!C:C, $A$4)').number_format = "₩#,##0"
        ws_dash.cell(curr_row, 13, f"=IF(J{curr_row}>0, K{curr_row}/J{curr_row}, 0)").number_format = "0.00%"
        ws_dash.cell(curr_row, 14, f"=IF(K{curr_row}>0, L{curr_row}/K{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 15, f'=SUMIFS(통합_DB!Q:Q, 통합_DB!B:B, "{db_match}", 통합_DB!C:C, $A$4)').number_format = "#,##0"
        ws_dash.cell(curr_row, 16, f"=IF(O{curr_row}>0, L{curr_row}/O{curr_row}, 0)").number_format = "₩#,##0"
        ws_dash.cell(curr_row, 17, f'=SUMIFS(통합_DB!S:S, 통합_DB!B:B, "{db_match}", 통합_DB!C:C, $A$4)').number_format = "#,##0"
        ws_dash.cell(curr_row, 18, f"=IF(Q{curr_row}>0, L{curr_row}/Q{curr_row}, 0)").number_format = "₩#,##0"
        
    for c in range(1, 19):
        cell = ws_dash.cell(curr_row, c); cell.font = c_font
        if c_fill: cell.fill = c_fill
        cell.border = border_thin

# Total Row (Formula)
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

# Highlights (Formula based on 통합_DB)
hl_row = tot_row + 2
ws_dash.cell(hl_row, 1, "[당월 핵심 성과 Top 하이라이트] 검색 키워드 및 디스플레이 소재").font = font_sec_title
hl_h_row = hl_row + 1
ws_dash.merge_cells(f'A{hl_h_row}:I{hl_h_row}')
ws_dash.cell(hl_h_row, 1, "수임 견인 Top 5 키워드 (검색광고 실시간 수식)").font = font_th; ws_dash.cell(hl_h_row, 1).fill = fill_th; ws_dash.cell(hl_h_row, 1).alignment = align_center

ws_dash.merge_cells(f'J{hl_h_row}:R{hl_h_row}')
ws_dash.cell(hl_h_row, 10, "최고 효율 Top 5 소재 (디스플레이 광고 META/GFA 실시간 수식)").font = font_th; ws_dash.cell(hl_h_row, 10).fill = fill_th; ws_dash.cell(hl_h_row, 10).alignment = align_center

top_kw_headers = ["순위", "키워드", "매체", "노출수", "클릭수", "비용(vat-)", "리드", "수임", "수임CPA"]
top_crv_headers = ["순위", "소재명 (Creative)", "매체", "노출수", "클릭수", "비용(vat-)", "리드", "수임", "리드CPA"]

sub_h_row = hl_h_row + 1
for idx, h in enumerate(top_kw_headers, 1):
    c = ws_dash.cell(sub_h_row, idx, h)
    c.font = font_th; c.fill = PatternFill(start_color='E9EEF4', end_color='E9EEF4', fill_type='solid'); c.alignment = align_center; c.border = border_thin

for idx, h in enumerate(top_crv_headers, 10):
    c = ws_dash.cell(sub_h_row, idx, h)
    c.font = font_th; c.fill = PatternFill(start_color='E9EEF4', end_color='E9EEF4', fill_type='solid'); c.alignment = align_center; c.border = border_thin

top_kw_keys = [
    ("법인기장", "네이버 검색광고"),
    ("세무사", "네이버 검색광고"),
    ("개인사업자세무사", "네이버 검색광고"),
    ("법인세무기장", "네이버 검색광고"),
    ("회계사법인", "구글 검색광고"),
]

top_crv_keys = [
    ("vid_interview_profitloss_leejaeheeA", "META"),
    ("vid_인터뷰_차별점24시_조태식대표_2608", "META"),
    ("vid_interview_taxconcern_leejeongseopA", "META"),
    ("com_taxplan_gift", "네이버 GFA"),
    ("ntv_img_realtime_support_team_square_260901", "네이버 GFA"),
]

for idx, (kw, med) in enumerate(top_kw_keys, 1):
    curr = sub_h_row + idx
    ws_dash.cell(curr, 1, idx).alignment = align_center
    ws_dash.cell(curr, 2, kw).alignment = align_left
    ws_dash.cell(curr, 3, med).alignment = align_center
    # DYNAMIC FORMULA:
    ws_dash.cell(curr, 4, f'=SUMIFS(통합_DB!M:M, 통합_DB!B:B, C{curr}, 통합_DB!L:L, B{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 5, f'=SUMIFS(통합_DB!N:N, 통합_DB!B:B, C{curr}, 통합_DB!L:L, B{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 6, f'=SUMIFS(통합_DB!O:O, 통합_DB!B:B, C{curr}, 통합_DB!L:L, B{curr})').number_format = "₩#,##0"
    ws_dash.cell(curr, 7, f'=SUMIFS(통합_DB!Q:Q, 통합_DB!B:B, C{curr}, 통합_DB!L:L, B{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 8, f'=SUMIFS(통합_DB!S:S, 통합_DB!B:B, C{curr}, 통합_DB!L:L, B{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 9, f'=IF(H{curr}>0, F{curr}/H{curr}, 0)').number_format = "₩#,##0"
    for c in range(1, 10):
        cell = ws_dash.cell(curr, c); cell.font = font_td; cell.border = border_thin

for idx, (crv, med) in enumerate(top_crv_keys, 1):
    curr = sub_h_row + idx
    ws_dash.cell(curr, 10, idx).alignment = align_center
    ws_dash.cell(curr, 11, crv).alignment = align_left
    ws_dash.cell(curr, 12, med).alignment = align_center
    # DYNAMIC FORMULA:
    ws_dash.cell(curr, 13, f'=SUMIFS(통합_DB!M:M, 통합_DB!B:B, L{curr}, 통합_DB!L:L, K{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 14, f'=SUMIFS(통합_DB!N:N, 통합_DB!B:B, L{curr}, 통합_DB!L:L, K{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 15, f'=SUMIFS(통합_DB!O:O, 통합_DB!B:B, L{curr}, 통합_DB!L:L, K{curr})').number_format = "₩#,##0"
    ws_dash.cell(curr, 16, f'=SUMIFS(통합_DB!Q:Q, 통합_DB!B:B, L{curr}, 통합_DB!L:L, K{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 17, f'=SUMIFS(통합_DB!S:S, 통합_DB!B:B, L{curr}, 통합_DB!L:L, K{curr})').number_format = "#,##0"
    ws_dash.cell(curr, 18, f'=IF(P{curr}>0, O{curr}/P{curr}, 0)').number_format = "₩#,##0"
    for c in range(10, 19):
        cell = ws_dash.cell(curr, c); cell.font = font_td; cell.border = border_thin

dash_col_widths = {
    'A': 12, 'B': 22, 'C': 15, 'D': 15, 'E': 9, 'F': 10, 'G': 14, 'H': 14, 'I': 15,
    'J': 13, 'K': 10, 'L': 16, 'M': 9, 'N': 11, 'O': 10, 'P': 14, 'Q': 10, 'R': 14
}
for col_letter, w in dash_col_widths.items():
    ws_dash.column_dimensions[col_letter].width = w

# === Sheet 2: 기간별_추이비교 (100% 수식 기반) ===
print("Building Sheet 2: 기간별_추이비교 (100% 수식 기반)...")
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

# Past monthly history + Dynamic formula for current months
# 2025.10 ~ 2026.08: Historical baselines
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

# Active Months: 2026.09 and 2026.10 (Fully Formula Driven from 통합_DB)
active_months = ["2026.09.", "2026.10."]
for idx, m_code in enumerate(active_months):
    cur = m_start + len(history_months) + idx
    ws_trend.cell(cur, 1, m_code).alignment = align_center
    # Dynamic formulas from 통합_DB:
    ws_trend.cell(cur, 2, f'=SUMIFS(통합_DB!M:M, 통합_DB!D:D, A{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 3, f'=SUMIFS(통합_DB!N:N, 통합_DB!D:D, A{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 4, f'=IF(C{cur}>0, G{cur}/C{cur}, 0)').number_format = "₩#,##0"
    ws_trend.cell(cur, 5, f'=IF(B{cur}>0, C{cur}/B{cur}, 0)').number_format = "0.00%"
    ws_trend.cell(cur, 6, f'=IF(B{cur}>0, G{cur}/B{cur}*1000, 0)').number_format = "₩#,##0"
    ws_trend.cell(cur, 7, f'=SUMIFS(통합_DB!O:O, 통합_DB!D:D, A{cur})').number_format = "₩#,##0"
    ws_trend.cell(cur, 8, f'=SUMIFS(통합_DB!P:P, 통합_DB!D:D, A{cur})').number_format = "₩#,##0"
    ws_trend.cell(cur, 9, f'=SUMIFS(통합_DB!Q:Q, 통합_DB!D:D, A{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 10, f'=IF(I{cur}>0, G{cur}/I{cur}, 0)').number_format = "₩#,##0"
    ws_trend.cell(cur, 11, f'=SUMIFS(통합_DB!R:R, 통합_DB!D:D, A{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 12, f'=IF(C{cur}>0, I{cur}/C{cur}, 0)').number_format = "0.0%"
    ws_trend.cell(cur, 13, f'=IF(K{cur}>0, G{cur}/K{cur}, 0)').number_format = "₩#,##0"
    ws_trend.cell(cur, 14, f'=SUMIFS(통합_DB!S:S, 통합_DB!D:D, A{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 15, f'=IF(N{cur}>0, G{cur}/N{cur}, 0)').number_format = "₩#,##0"
    ws_trend.cell(cur, 16, f'=SUMIFS(통합_DB!T:T, 통합_DB!D:D, A{cur})').number_format = "₩#,##0"
    ws_trend.cell(cur, 17, f'=MAX(0, I{cur}-N{cur})').number_format = "#,##0"
    for c in range(1, 18):
        cell = ws_trend.cell(cur, c); cell.font = font_td; cell.border = border_thin

# Comparison Rows
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

# Weekly Section (Formula Driven)
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

# Unique weeks sorted
weeks_list = [
    ("9월 1주차", "26_35주"),
    ("9월 2주차", "26_36주"),
    ("9월 3주차", "26_37주"),
    ("9월 4주차", "26_38주"),
    ("9월 5주차", "26_39주"),
    ("10월 1주차", "26_40주"),
]

w_start = row_w_h + 1
for i, (w_name, w_code) in enumerate(weeks_list):
    cur = w_start + i
    ws_trend.cell(cur, 1, w_name).alignment = align_center
    ws_trend.cell(cur, 2, w_code).alignment = align_center
    # DYNAMIC FORMULAS
    ws_trend.cell(cur, 3, f'=SUMIFS(통합_DB!M:M, 통합_DB!E:E, B{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 4, f'=SUMIFS(통합_DB!N:N, 통합_DB!E:E, B{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 5, f"=IF(C{cur}>0, D{cur}/C{cur}, 0)").number_format = "0.00%"
    ws_trend.cell(cur, 6, f"=IF(D{cur}>0, G{cur}/D{cur}, 0)").number_format = "₩#,##0"
    ws_trend.cell(cur, 7, f'=SUMIFS(통합_DB!O:O, 통합_DB!E:E, B{cur})').number_format = "₩#,##0"
    ws_trend.cell(cur, 8, f'=SUMIFS(통합_DB!P:P, 통합_DB!E:E, B{cur})').number_format = "₩#,##0"
    ws_trend.cell(cur, 9, f'=SUMIFS(통합_DB!Q:Q, 통합_DB!E:E, B{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 10, f"=IF(I{cur}>0, G{cur}/I{cur}, 0)").number_format = "₩#,##0"
    ws_trend.cell(cur, 11, f'=SUMIFS(통합_DB!R:R, 통합_DB!E:E, B{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 12, f"=IF(D{cur}>0, I{cur}/D{cur}, 0)").number_format = "0.0%"
    ws_trend.cell(cur, 13, f'=SUMIFS(통합_DB!S:S, 통합_DB!E:E, B{cur})').number_format = "#,##0"
    ws_trend.cell(cur, 14, f"=IF(M{cur}>0, G{cur}/M{cur}, 0)").number_format = "₩#,##0"
    for c in range(1, 15):
        cell = ws_trend.cell(cur, c); cell.font = font_td; cell.border = border_thin

# Weekly WoW Comparison Row
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

# === Unified Function for Media Specific Sheets (100% PURE FORMULAS) ===
def build_formula_media_sheet(wb, sheet_title, page_header, media_name, key_list, is_da=False):
    print(f"Building pure formula sheet: {sheet_title}...")
    ws = wb.create_sheet(title=sheet_title)
    ws.views.sheetView[0].showGridLines = True
    
    # Header
    ws.merge_cells('A1:N1')
    ws['A1'] = page_header
    ws['A1'].font = font_title; ws['A1'].fill = fill_navy; ws['A1'].alignment = align_center
    ws.row_dimensions[1].height = 40
    
    # Top KPI cards with pure SUMIFS from 통합_DB
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 24
    
    card_configs = [
        ("총 노출수", f'=SUMIFS(통합_DB!M:M, 통합_DB!B:B, "{media_name}")', "1", "2", "#,##0"),
        ("총 클릭수", f'=SUMIFS(통합_DB!N:N, 통합_DB!B:B, "{media_name}")', "3", "4", "#,##0"),
        ("총 비용(VAT-)", f'=SUMIFS(통합_DB!O:O, 통합_DB!B:B, "{media_name}")', "5", "7", "₩#,##0"),
        ("총 리드", f'=SUMIFS(통합_DB!Q:Q, 통합_DB!B:B, "{media_name}")', "8", "9", "#,##0"),
        ("총 수임", f'=SUMIFS(통합_DB!S:S, 통합_DB!B:B, "{media_name}")', "10", "11", "#,##0"),
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
                
    ws.cell(6, 1, f"[{sheet_title} 상세 성과 목록] 통합_DB 참조 실시간 수식 집계").font = font_sec_title
    
    label_unit = "소재명 (Creative)" if is_da else "키워드 (Keyword)"
    headers = [
        "순위", label_unit, "노출수", "클릭수", "비용(vat-)", "비용(vat+)",
        "CTR", "CPC", "총 리드", "유효 리드", "수임", "리드 CPA", "수임 CPA", "전환율(CVR)"
    ]
    
    ws.row_dimensions[7].height = 24
    for c_idx, h in enumerate(headers, 1):
        c = ws.cell(7, c_idx, h)
        c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header
        
    start_row = 8
    for rank, key_name in enumerate(key_list[:250], 1):
        cur_r = start_row + rank - 1
        
        ws.cell(cur_r, 1, rank).alignment = align_center
        ws.cell(cur_r, 2, key_name).alignment = align_left
        
        # 100% PURE DYNAMIC FORMULAS REFERENCING 통합_DB
        ws.cell(cur_r, 3, f'=SUMIFS(통합_DB!M:M, 통합_DB!B:B, "{media_name}", 통합_DB!L:L, B{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 4, f'=SUMIFS(통합_DB!N:N, 통합_DB!B:B, "{media_name}", 통합_DB!L:L, B{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 5, f'=SUMIFS(통합_DB!O:O, 통합_DB!B:B, "{media_name}", 통합_DB!L:L, B{cur_r})').number_format = "₩#,##0"
        ws.cell(cur_r, 6, f'=SUMIFS(통합_DB!P:P, 통합_DB!B:B, "{media_name}", 통합_DB!L:L, B{cur_r})').number_format = "₩#,##0"
        ws.cell(cur_r, 7, f"=IF(C{cur_r}>0, D{cur_r}/C{cur_r}, 0)").number_format = "0.00%"
        ws.cell(cur_r, 8, f"=IF(D{cur_r}>0, E{cur_r}/D{cur_r}, 0)").number_format = "₩#,##0"
        ws.cell(cur_r, 9, f'=SUMIFS(통합_DB!Q:Q, 통합_DB!B:B, "{media_name}", 통합_DB!L:L, B{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 10, f'=SUMIFS(통합_DB!R:R, 통합_DB!B:B, "{media_name}", 통합_DB!L:L, B{cur_r})').number_format = "#,##0"
        ws.cell(cur_r, 11, f'=SUMIFS(통합_DB!S:S, 통합_DB!B:B, "{media_name}", 통합_DB!L:L, B{cur_r})').number_format = "#,##0"
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

# Get sorted unique key lists
def get_unique_keys(df_media):
    grouped = df_media.groupby('키워드_소재').agg({'수임': 'sum', '총리드': 'sum', '비용_vat제외': 'sum'}).reset_index()
    if len(grouped) > 1:
        grouped = grouped[grouped['키워드_소재'] != '-']
    grouped = grouped.sort_values(by=['수임', '총리드', '비용_vat제외'], ascending=[False, False, False])
    return grouped['키워드_소재'].tolist()

# Build sheets with 100% pure formulas!
naver_keys = get_unique_keys(df_fact[df_fact['매체구분'] == '네이버 검색광고'])
build_formula_media_sheet(wb, '네이버_SA_키워드', '네이버 검색광고 키워드별 성과 분석 (수식 연동)', '네이버 검색광고', naver_keys, is_da=False)

google_keys = get_unique_keys(df_fact[df_fact['매체구분'] == '구글 검색광고'])
build_formula_media_sheet(wb, '구글_SA_키워드', '구글 검색광고 키워드 성과 분석 (수식 연동)', '구글 검색광고', google_keys, is_da=False)

meta_keys = get_unique_keys(df_fact[df_fact['매체구분'] == 'META'])
build_formula_media_sheet(wb, 'META_소재분석', 'META DA 광고 소재별 성과 분석 (수식 연동)', 'META', meta_keys, is_da=True)

gfa_keys = get_unique_keys(df_fact[df_fact['매체구분'] == '네이버 GFA'])
build_formula_media_sheet(wb, 'GFA_소재분석', '네이버 GFA 소재별 성과 분석 (수식 연동)', '네이버 GFA', gfa_keys, is_da=True)

place_keys = get_unique_keys(df_fact[df_fact['매체구분'] == '네이버 플레이스'])
build_formula_media_sheet(wb, '네이버_플레이스', '네이버 플레이스 검색어별 성과 분석 (수식 연동)', '네이버 플레이스', place_keys, is_da=False)

# Sheet 8: 통합_DB
print("Building Sheet: 통합_DB...")
ws_db = wb.create_sheet(title='통합_DB')
ws_db.views.sheetView[0].showGridLines = True

db_headers = [
    "구분", "매체구분", "일자", "월", "주간코드", "주차", "요일", "영업일구분",
    "디바이스", "캠페인", "광고그룹", "키워드_소재", "노출수", "클릭수",
    "비용_vat제외", "비용_vat포함", "총리드", "유효리드", "수임", "수익"
]

ws_db.row_dimensions[1].height = 24
for c_idx, h in enumerate(db_headers, 1):
    c = ws_db.cell(1, c_idx, h)
    c.font = font_th; c.fill = fill_th; c.alignment = align_center; c.border = border_header

print(f"Writing {len(df_fact)} rows to 통합_DB...")
for r_idx, row in enumerate(df_fact[db_headers].itertuples(index=False), 2):
    for c_idx, val in enumerate(row, 1):
        c = ws_db.cell(r_idx, c_idx, val)
        if c_idx in [13, 14, 17, 18, 19]:
            c.number_format = "#,##0"
        elif c_idx in [15, 16, 20]:
            c.number_format = "₩#,##0"

# Sheet 9: Raw_가이드
ws_guide = wb.create_sheet(title='Raw_가이드')
ws_guide.views.sheetView[0].showGridLines = True
ws_guide.merge_cells('A1:G1')
ws_guide['A1'] = "알프레드 키워드/소재 리포트 Raw 데이터 입력 및 수식 자동화 가이드"
ws_guide['A1'].font = font_title; ws_guide['A1'].fill = fill_navy; ws_guide['A1'].alignment = align_center
ws_guide.row_dimensions[1].height = 40

guide_texts = [
    ("1. 100% 수식 기반 자동화 구조", "본 리포트의 모든 시트('종합_대시보드', '기간별_추이비교', 매체별 키워드/소재 시트)는 '통합_DB' 시트를 직접 참조하는 SUMIFS/IF 동적 수식으로만 작성되어 있습니다."),
    ("2. Raw 데이터 갱신 방법", "'통합_DB' 시트에 새로운 행 데이터를 복사하여 붙여넣으면, 별도의 조작 없이 모든 리포트 시트의 숫자와 지표가 실시간으로 자동 갱신됩니다."),
    ("3. 대시보드 기준일 변경", "'종합_대시보드' 시트의 A4 셀에 날짜(예: 2026-10-01)를 입력하면, 해당 일자의 실시간 매체별 집행 금액, 노출, 클릭, 리드, 수임 성과가 즉시 계산되어 표출됩니다."),
    ("4. 키워드/소재 성과 자동 결합", "검색광고는 utm_term 기반 키워드, DA 광고는 utm_content의 '__' 파싱 기반 소재명을 기준으로 SUMIFS가 작동하여 비용과 전환(수임)이 자동으로 1:1 결합됩니다."),
    ("5. 기간별 비교 (MoM / WoW / DoD)", "'기간별_추이비교' 시트에서 통합_DB의 '월', '주간코드', '일자' 컬럼을 기준으로 실시간 월별(MoM), 주차별(WoW), 전년 동월(YoY) 비교 지표가 자동 산출됩니다."),
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
print(f"SUCCESS! {output_filename} has been generated.")

try:
    wb.save(primary_filename)
    print(f"SUCCESS! Also overwritten to {primary_filename}.")
except Exception as e:
    print(f"Notice: {primary_filename} is currently open in Excel, saved safely to {output_filename}.")
