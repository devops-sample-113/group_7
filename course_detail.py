from tkinter import *
import tkinter as tk
import pandas
import subprocess

##前往編輯課程資訊
def edit_course(code):
    
    teacher=number.get()
    # 清空輸入框
    number.delete(0, END)
    
    if all_course[code]['授課教師證號']==teacher:
        with open("course.txt","w") as file:
            file.write(str(code))
        subprocess.Popen(["python", "course_edit.py"])
    else:
        error_window=Tk()
        error_window.title("存取失敗")
        error_window.geometry('300x150')

        label_error = tk.Label(error_window, text="輸入證號與授課教師證號不符", fg="red",font=20)
        label_error.place(x=150,y=75,anchor="center")
        
def show_line(string):
    label = tk.Label(detail_window, text=string, font=("Arial", 10), anchor="w")
    label.pack(pady=10, padx=10, fill='x', anchor='w')

##課表搜尋頁面視窗初始設定
detail_window=Tk()
detail_window.title("課程資訊頁面")
detail_window.geometry('500x545')

try:
    with open("course.txt","r") as file:
        code=file.read()
except FileNotFoundError:
    code="尋找錯誤"

 # 讀取 Excel 文件
work = pandas.read_excel('資料庫.xlsx',sheet_name='課程')
# 以課程代碼作為 key，生成字典
all_course = work.set_index('課程代碼').T.to_dict()

##顯示課程資訊
code=int(code)
# 檢查 all_course 中是否存在該課程代碼
if code in all_course:
    # 從 all_course 取得課程名稱
    opt = all_course[code]['課程名稱']
else:
    opt = "找不到課程名稱"

#顯示課程名稱
string= f"{opt}"
label = tk.Label(detail_window, text=string, font=("Arial", 20, "bold"))
label.pack(pady=10)

#顯示其他課程資訊
string= f"課程代碼: {code}"
label = tk.Label(detail_window, text=string, font=("Arial", 10), anchor="w")
label.pack(pady=10, padx=10, fill='x', anchor='w')

opt=all_course[code]['開課時間']
show_line(f"開課時間: {opt}")

opt=all_course[code]['上課地點']
show_line(f"上課地點: {opt}")

opt=all_course[code]['授課教授']
show_line(f"授課教授: {opt}")

ta1=all_course[code]['課堂助教1']
ta2=all_course[code]['課堂助教2']
if pandas.notna(ta1):
    if pandas.notna(ta2):
        show_line(f"課堂助教: {ta1}、{ta2}")
    else:
        show_line(f"課堂助教: {ta1}")

opt=all_course[code]['課程大綱']
show_line(f"課程大綱: {opt}")

opt=all_course[code]['評分方式']
show_line(f"評分方式: {opt}")

opt=all_course[code]['修課人數上限']
show_line(f"修課人數上限: {opt}")

opt=all_course[code]['目前可修課人數餘額']
show_line(f"目前可修課人數餘額: {opt}")

opt=all_course[code]['學分']
show_line(f"學分: {opt}")

## 建立一個 frame 將輸入框與按鈕放入同一行
input_frame = tk.Frame(detail_window)
input_frame.pack(anchor="w", padx=10, pady=20, fill="x")

## 教師證號輸入框
number = tk.Entry(input_frame)
number.pack(side="left", padx=5)

## 搜尋按鈕
check = tk.Button(input_frame, text="編輯", command=lambda: edit_course(code))
check.pack(side="left", padx=5)

##將course.txt暫存資料清空
with open("course.txt","w") as file:
    file.write("")

detail_window.mainloop()