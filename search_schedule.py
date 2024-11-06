from tkinter import *
import tkinter as tk
import subprocess

##前往查詢
def show_search():
    #取得輸入資料並前往查詢
    with open("data.txt","w") as file:
        file.write(number.get())

    search_window.destroy()  # 關閉當前視窗
    subprocess.Popen(["python", "schedule.py"])  # 執行第二個程式

##課表搜尋頁面視窗初始設定
search_window=Tk()
search_window.title("個人課表查詢")
search_window.geometry('300x150')

##學號輸入框
number=tk.Entry(search_window)
number.pack(pady=20)

##搜尋按鈕
search = tk.Button(search_window, text="查詢", command=show_search)
search.pack(pady=20)

search_window.mainloop()