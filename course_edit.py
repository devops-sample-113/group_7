from tkinter import *
import tkinter as tk
import pandas
import openpyxl
import subprocess

def add_entry(label_text, initial_value, parent):
    tk.Label(parent, text=label_text).pack(anchor="w")
    entry = tk.Entry(parent)
    # 檢查是否為空值或 nan 值，若不是，才插入初始值
    if pandas.notna(initial_value):
        entry.insert(0, str(initial_value))
    entry.pack(fill="x", padx=10, pady=5)
    return entry

def save_changes():
    try:

        credits = entry_credits.get()
        if credits not in ["1", "2", "3"]:
            raise ValueError("學分只能為1、2或3，且不可為空")

        if not entry_course_name.get().strip():
            raise ValueError("課程名稱不可為空")
        if not entry_location.get().strip():
            raise ValueError("上課地點不可為空！")
        if not entry_outline.get().strip():
            raise ValueError("課程大綱不可為空！")
        if not entry_evaluation.get().strip():
            raise ValueError("評分方式不可為空！")

            # 檢查「修課人數上限」是否為正整數
        max_students = entry_max_students.get().strip()
        if not max_students.isdigit() or int(max_students) <= 0:
            raise ValueError("修課人數上限需為正整數，且不可為空")

        # 檢查「目前可修課人數餘額」是否為不為負的整數，且不能超過修課人數上限
        available_spots = entry_available_spots.get().strip()
        if not available_spots.isdigit() or int(available_spots) < 0:
            raise ValueError("目前可修課人數餘額需為非負整數，且不可為空")
        if int(available_spots) > int(max_students):
            raise ValueError("目前可修課人數餘額不能超過修課人數上限")
        
        # 更新 all_course 字典中可修改的屬性
        all_course[code]['課程名稱'] = entry_course_name.get()
        all_course[code]['上課地點'] = entry_location.get()
        all_course[code]['課堂助教1'] = entry_ta1.get()
        all_course[code]['課堂助教2'] = entry_ta2.get()
        all_course[code]['課程大綱'] = entry_outline.get()
        all_course[code]['評分方式'] = entry_evaluation.get()
        all_course[code]['修課人數上限'] = entry_max_students.get()
        all_course[code]['目前可修課人數餘額'] = entry_available_spots.get()
        all_course[code]['學分'] = entry_credits.get()

        # 儲存更改到 Excel 文件
        for key, value in all_course[code].items():
            # 檢查 DataFrame 的欄位類型並進行相應轉型
            if work[key].dtype == 'int64':  # 如果欄位需要整數
                all_course[code][key] = int(value)
            elif work[key].dtype == 'float64':  # 如果欄位需要浮點數
                all_course[code][key] = float(value)
            else:  # 否則，轉換為字串
                all_course[code][key] = str(value)

        # 更新 DataFrame
        work.loc[work['課程代碼'] == code, list(all_course[code].keys())] = list(all_course[code].values())
        # 儲存更改到 Excel 文件
        with pandas.ExcelWriter('資料庫.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            work.to_excel(writer, sheet_name='課程', index=False)

        # 顯示成功訊息視窗並關閉編輯頁面
        success_window = tk.Toplevel(detail_window)
        success_window.title("編輯")
        success_window.geometry('300x150')
        label_success = tk.Label(success_window, text=f"{code}課程資訊編輯成功", fg="blue", font=20)
        label_success.place(x=150, y=75, anchor="center")
        detail_window.destroy()

    except ValueError as ve:
        # 顯示課程名稱為空的錯誤
        error_window = tk.Toplevel(detail_window)
        error_window.title("保存失敗")
        error_window.geometry('500x150')
        label_error = tk.Label(error_window, text=f"保存失敗: {str(ve)}", fg="red", font=20)
        label_error.place(x=250, y=75, anchor="center")

    except Exception as e:
        error_window = tk.Toplevel(detail_window)
        error_window.title("編輯失敗")
        error_window.geometry('300x150')
        label_error = tk.Label(error_window, text=f"編輯失敗: {str(e)}", fg="red", font=20)
        label_error.place(x=150, y=75, anchor="center")

## 課程資訊頁面
detail_window = Tk()
detail_window.title("課程編輯頁面")
detail_window.geometry('500x650')

# 建立一個頂部的 Frame 來放置按鈕，並靠右對齊
top_frame = tk.Frame(detail_window)
top_frame.pack(fill="x", side="top", anchor="ne")  # 設定頂部對齊並充滿整個寬度

# 在 Frame 中放置保存按鈕，並靠右對齊
save_button = tk.Button(top_frame, text="保存", command=save_changes)
save_button.pack(side="right", padx=10, pady=10)  # 右側對齊並設定邊距

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
    tk.Label(detail_window, text=f"課程代碼: {code}").pack(anchor="w", pady=5)
    tk.Label(detail_window, text=f"開課時間: {all_course[code]['開課時間']}").pack(anchor="w", pady=5)
    tk.Label(detail_window, text=f"授課教授: {all_course[code]['授課教授']}").pack(anchor="w", pady=5)
    tk.Label(detail_window, text=f"教師證號: {all_course[code]['授課教師證號']}").pack(anchor="w", pady=5)

    ta1_id = all_course[code]['課堂助教1'] if pandas.notna(all_course[code]['課堂助教1']) else ""
    ta2_id = all_course[code]['課堂助教2'] if pandas.notna(all_course[code]['課堂助教2']) else ""

    #可更改課程資訊
    entry_course_name = add_entry("課程名稱:", all_course[code]['課程名稱'], detail_window)
    # 其他課程資訊
    entry_location = add_entry("上課地點:", all_course[code]['上課地點'], detail_window)
    entry_ta1 = add_entry("課堂助教1證號:", ta1_id, detail_window)
    entry_ta2 = add_entry("課堂助教2證號:", ta2_id, detail_window)
    entry_outline = add_entry("課程大綱:", all_course[code]['課程大綱'], detail_window)
    entry_evaluation = add_entry("評分方式:", all_course[code]['評分方式'], detail_window)
    entry_max_students = add_entry("修課人數上限:", all_course[code]['修課人數上限'], detail_window)
    entry_available_spots = add_entry("目前可修課人數餘額:", all_course[code]['目前可修課人數餘額'], detail_window)
    entry_credits = add_entry("學分:", all_course[code]['學分'], detail_window)
else:
    tk.Label(detail_window, text="找不到課程名稱").pack()

detail_window.mainloop()
