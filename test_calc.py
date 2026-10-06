# -*- coding: utf-8 -*-
import openpyxl

wb = openpyxl.Workbook()
ws1 = wb.active
ws1.title = 'raw'
ws1.append(['val', 'code', 'date'])
ws1.append([100, '26_38주(09/21~09/27)', '2026-09-21'])
ws1.append([200, '26_38주(09/21~09/27)', '2026-09-22'])
ws1.append([300, '26_39주(09/28~10/04)', '2026-09-28'])

ws2 = wb.create_sheet('calc')
ws2.append(['code', 'sum_wildcard', 'sum_exact'])
ws2.append(['26_38주', '=SUMIFS(raw!A:A, raw!B:B, A2 & "*")', '=SUMIFS(raw!A:A, raw!B:B, A2)'])
ws2.append(['26_38주(09/21~09/27)', '=SUMIFS(raw!A:A, raw!B:B, A3 & "*")', '=SUMIFS(raw!A:A, raw!B:B, A3)'])

wb.save('test_wildcard.xlsx')
print('Successfully saved test_wildcard.xlsx')
