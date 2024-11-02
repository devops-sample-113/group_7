from tkinter import *
import tkinter as tk
import openpyxl
import subprocess

##前往查詢
def show_search():
    search_window.destroy()  # 關閉當前視窗
    subprocess.Popen(["python", "schedule.py"])  # 執行第二個程式

##課表搜尋頁面視窗初始設定
search_window=Tk()
search_window.title("個人課表查詢")
search_window.geometry('500x400+390+75')

##搜尋按鈕
search = tk.Button(search_window, text="查詢", command=show_search)
search.pack(pady=20)

search_window.mainloop()
