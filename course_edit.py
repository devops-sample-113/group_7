from tkinter import *
import tkinter as tk
import pandas
import openpyxl
import subprocess

def save_changes():
    try:
        # 更新 all_course 字典中可修改的屬性
        all_course[code]['課程名稱'] = entry_course_name.get()
        all_course[code]['開課時間'] = entry_time.get()
        all_course[code]['上課地點'] = entry_location.get()
        all_course[code]['課堂助教1'] = entry_ta1.get()
        all_course[code]['課堂助教2'] = entry_ta2.get()
        all_course[code]['課程大綱'] = entry_outline.get()
        all_course[code]['評分方式'] = entry_evaluation.get()
        all_course[code]['修課人數上限'] = entry_max_students.get()
        all_course[code]['目前可修課人數餘額'] = entry_available_spots.get()
        all_course[code]['學分'] = entry_credits.get()

        # 儲存更改到 Excel 文件
        work.loc[work['課程代碼'] == code, list(all_course[code].keys())] = list(all_course[code].values())
        with pandas.ExcelWriter('資料庫.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            work.to_excel(writer, sheet_name='課程', index=False)

        # 顯示成功訊息視窗並關閉編輯頁面
        success_window = tk.Toplevel(detail_window)
        success_window.title("編輯")
        success_window.geometry('300x150')
        label_success = tk.Label(success_window, text=f"{code}課程資訊編輯成功", fg="blue", font=20)
        label_success.place(x=150, y=75, anchor="center")
        detail_window.destroy()

    except Exception as e:
        error_window = tk.Toplevel(detail_window)
        error_window.title("編輯失敗")
        error_window.geometry('300x150')
        label_error = tk.Label(error_window, text=f"編輯失敗: {str(e)}", fg="red", font=20)
        label_error.place(x=150, y=75, anchor="center")

## 課程資訊頁面
detail_window = Tk()
detail_window.title("課程編輯頁面")
detail_window.geometry('500x600')

try:
    with open("course.txt", "r") as file:
        code = int(file.read())
except FileNotFoundError:
    code = "尋找錯誤"

# 清空 course.txt
with open("course.txt", "w") as file:
    file.write("")

# 讀取 Excel 文件
work = pandas.read_excel('資料庫.xlsx', sheet_name='課程')
all_course = work.set_index('課程代碼').T.to_dict()

# 顯示課程名稱標題
tk.Label(detail_window, text="課程編輯", font=("Arial", 20, "bold")).pack(pady=10)

# 設置輸入框和初始值
def add_entry(label_text, initial_value, parent):
    tk.Label(parent, text=label_text).pack(anchor="w")
    entry = tk.Entry(parent)
    entry.insert(0, str(initial_value))
    entry.pack(fill="x", padx=10, pady=5)
    return entry

# 顯示並編輯課程資訊
if code in all_course:
    entry_course_name = add_entry("課程名稱:", all_course[code]['課程名稱'], detail_window)
    tk.Label(detail_window, text=f"課程代碼: {code}").pack(anchor="w", pady=5)
    entry_time = add_entry("開課時間:", all_course[code]['開課時間'], detail_window)
    entry_location = add_entry("上課地點:", all_course[code]['上課地點'], detail_window)
    entry_ta1 = add_entry("課堂助教1:", all_course[code]['課堂助教1'], detail_window)
    entry_ta2 = add_entry("課堂助教2:", all_course[code]['課堂助教2'], detail_window)
    entry_outline = add_entry("課程大綱:", all_course[code]['課程大綱'], detail_window)
    entry_evaluation = add_entry("評分方式:", all_course[code]['評分方式'], detail_window)
    entry_max_students = add_entry("修課人數上限:", all_course[code]['修課人數上限'], detail_window)
    entry_available_spots = add_entry("目前可修課人數餘額:", all_course[code]['目前可修課人數餘額'], detail_window)
    entry_credits = add_entry("學分:", all_course[code]['學分'], detail_window)
else:
    tk.Label(detail_window, text="找不到課程名稱").pack()

# 保存按鈕
save_button = tk.Button(detail_window, text="保存", command=save_changes)
save_button.pack(pady=20)

detail_window.mainloop()
