from tkinter import *
import tkinter as tk
import pandas
import openpyxl
import subprocess

def save_changes():
    try:

        # 檢查「課程名稱」是否為空
        if not entry_course_name.get().strip():
            raise ValueError("課程名稱不可為空")
        if not entry_location.get().strip():
            raise ValueError("上課地點不可為空！")
        
        # 格式化開課時間
        course_day = day_var.get()
        course_start_time = start_time_var.get()
        course_end_time = end_time_var.get()
        course_time = f"{course_day} {course_start_time}-{course_end_time}"
        
        # 更新 all_course 字典中可修改的屬性
        all_course[code]['課程名稱'] = entry_course_name.get()
        all_course[code]['開課時間'] = course_time
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
        error_window.geometry('300x150')
        label_error = tk.Label(error_window, text=f"保存失敗: {str(ve)}", fg="red", font=20)
        label_error.place(x=150, y=75, anchor="center")

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
    tk.Label(detail_window, text=f"授課教授: {all_course[code]['授課教授']}").pack(anchor="w", pady=5)
    tk.Label(detail_window, text=f"教師證號: {all_course[code]['授課教師證號']}").pack(anchor="w", pady=5)
    entry_course_name = add_entry("課程名稱:", all_course[code]['課程名稱'], detail_window)
    
    # 開課時間標籤
    tk.Label(detail_window, text="開課時間:").pack(anchor="w")

    # 建立開課時間的容器框架
    time_frame = Frame(detail_window)
    time_frame.pack(fill="x", padx=10, pady=5)

    if '開課時間' in all_course[code] and all_course[code]['開課時間']:
        course_day, course_time = all_course[code]['開課時間'].split()
        course_start_time, course_end_time = course_time.split('-')
    else:
        # 預設值，如果無法取得開課時間
        course_day, course_start_time, course_end_time = "星期一", "08:00", "09:00"

    # 星期選項
    day_var = StringVar()
    day_var.set(course_day)
    day_options = ["星期一", "星期二", "星期三", "星期四", "星期五"]
    day_menu = OptionMenu(time_frame, day_var, *day_options)
    day_menu.pack(side="left")

    # 課程開始時間選項
    start_time_var = StringVar()
    start_time_var.set(course_start_time)
    start_time_options = ["08:00", "09:00", "10:00", "11:00", "13:00", "14:00", "15:00", "16:00"]
    start_time_menu = OptionMenu(time_frame, start_time_var, *start_time_options)
    start_time_menu.pack(side="left")

    # 加入 "-" 標記
    separator_label = Label(time_frame, text="~")
    separator_label.pack(side="left")

    # 課程結束時間選項
    end_time_var = StringVar()
    end_time_var.set(course_end_time)
    end_time_options = ["09:00", "10:00", "11:00", "12:00", "14:00", "15:00", "16:00", "17:00"]
    end_time_menu = OptionMenu(time_frame, end_time_var, *end_time_options)
    end_time_menu.pack(side="left")
    
    # 其他課程資訊
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
