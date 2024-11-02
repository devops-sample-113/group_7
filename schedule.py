from tkinter import *
import tkinter as tk
import openpyxl
import subprocess

##返回課表搜尋頁面
def return_search():
    Swindow.destroy()  # 關閉當前視窗
    subprocess.Popen(["python", "search_schedule.py"])  # 執行第二個程式

##課表頁面視窗初始設定
Swindow=Tk()
Swindow.title("你的個人課表")
Swindow.geometry('500x400+390+75')

##個人課表顯示
#開啟excel
path='C:\\Users\\User\\Documents\\GitHub\\group_7\\個人課表\\D1234567.xlsx'
workbook=openpyxl.load_workbook(path)
#選擇工作表
worksheet=workbook.active

# 讀取 1-10 行，A-F 列的內容並顯示在 Label 中
for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=10, min_col=1, max_col=6, values_only=True)):
    for col_idx, value in enumerate(row):
        label = tk.Label(Swindow, text=value if value is not None else "", borderwidth=1, relief="solid", padx=5, pady=5)
        label.grid(row=row_idx, column=col_idx, sticky="nsew", padx=2, pady=2)

# 自動調整列寬
for col in range(6):
    Swindow.grid_columnconfigure(col, weight=1)

##返回按鈕
back=Button(Swindow,text="返回",anchor="s",command=return_search)
back.grid(row=20,column=5,pady=10)

Swindow.mainloop()