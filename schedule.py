from tkinter import *
import tkinter as tk
import openpyxl

##課表頁面視窗初始設定
Swindow=Tk()
Swindow.title("個人課表查詢")
Swindow.geometry('500x500+390+75')

#創建空label以存入課表以顯示
headers = ["A", "B", "C", "D", "E", "F"]
for col, header in enumerate(headers):
    label = tk.Label(Swindow, text=header, font=("Arial", 10, "bold"), borderwidth=1, relief="solid")
    label.grid(row=0, column=col, sticky="nsew", padx=2, pady=2)

##個人課表顯示
#開啟excel
path='/Users/hayashitogi/Documents/GitHub/group_7/資料庫.xlsx'
workbook=openpyxl.load_workbook(path)
#選擇工作表
worksheet=workbook.active

# 讀取 1-10 行，A-F 列的內容並顯示在 Label 中
for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=10, min_col=1, max_col=6, values_only=True), start=1):
    for col_idx, value in enumerate(row):
        label = tk.Label(Swindow, text=value if value is not None else "", borderwidth=1, relief="solid", padx=5, pady=5)
        label.grid(row=row_idx, column=col_idx, sticky="nsew", padx=2, pady=2)

# 自動調整列寬
for col in range(6):
    Swindow.grid_columnconfigure(col, weight=1)

Swindow.mainloop()