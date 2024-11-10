from tkinter import *
from tkinter import messagebox
import tkinter as tk
import re
import subprocess

##前往查詢
def show_search():

    input_text = number.get()
    
    # 檢查格式是否符合指定要求
    if re.match(r"^[DTA]\d{7}$", input_text):
        # 如果格式正確，將輸入資料寫入檔案並執行下一步
        with open("data.txt", "w") as file:
            file.write(input_text)
        search_window.destroy()  # 關閉當前視窗
        subprocess.Popen(["python", "schedule_show.py"])  # 執行第二個程式
    else:
        # 格式不符合，彈出錯誤訊息
        messagebox.showerror("輸入錯誤", "學號(證號)格式錯誤")
        number.delete(0, tk.END)

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