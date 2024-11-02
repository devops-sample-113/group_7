import os
os.chdir('/Users/hayashitogi/Documents/GitHub/group_7/資料庫.xlsx')

import openpyxl
wb = openpyxl.load_workbook('資料庫.xlsx', data_only=True)  # 設定 data_only=True 只讀取計算後的數值

s1 = wb['課程']
v = s1.iter_rows(min_row=1, min_col=1, max_col=53, max_row=5)  # 取出四格內容
print(v)
for i in v:
    for j in i:
        print(j.value)