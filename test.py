from tkinter import *
import tkinter as tk
import openpyxl
from tkinter import messagebox

Swindow = Tk()
Swindow.title("選課系統")
Swindow.geometry("1500x1000")

canvas = Canvas(Swindow, highlightthickness=0)
scrollbar = Scrollbar(Swindow, orient="vertical", command=canvas.yview)
scrollbar.pack(side="right", fill="y")
canvas.configure(yscrollcommand=scrollbar.set)
canvas.place(x=10, y=50, width=1480, height=700)

content_frame = Frame(canvas)
canvas.create_window((0, 0), window=content_frame, anchor="nw")

def on_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

content_frame.bind("<Configure>", on_configure)

# Adjust column widths
def adjust_column_widths(event=None):
    canvas_width = canvas.winfo_width()
    col_width = canvas_width // len(headers)
    for col, header in enumerate(headers):
        label = tk.Label(content_frame, text=header, font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=col_width // 10)
        label.grid(row=0, column=col, sticky="nsew", padx=2, pady=2)

canvas.bind("<Configure>", adjust_column_widths)

# Mouse wheel scrolling
def on_mouse_wheel(event):
    if event.delta > 0:
        canvas.yview_scroll(-1, "units")
    elif event.delta < 0:
        canvas.yview_scroll(1, "units")

canvas.bind_all("<MouseWheel>", on_mouse_wheel)
canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

def open_new_window():
    # 創建新的 Toplevel 視窗
    new_window = Toplevel(Swindow)
    new_window.title("課表頁面")
    new_window.geometry("400x300")  # 設置新視窗的大小

    # 在新視窗上添加一些元件
    label = Label(new_window, text="個人課表")
    label.pack(pady=20)

    close_button = Button(new_window, text="關閉", command=new_window.destroy)
    close_button.pack(pady=10)

# 按鍵，點擊後開啟新頁面
button_pop = Button(Swindow, text="課表頁面", command=open_new_window)
button_pop.place(x=10, y=10)

# Excel paths
database_path = '/Users/hayashitogi/Documents/GitHub/group_7/資料庫.xlsx'
personal_schedule_path = '/Users/hayashitogi/Documents/GitHub/group_7/個人課表.xlsx'

# Open the course database Excel
workbook = openpyxl.load_workbook(database_path)
worksheet = workbook["課程"]

# Define headers and the mapping for time slots and days
headers = ["課程名稱", "課程代碼", "開課時間", "上課地點", "授課教授", "加退選匡"]

# Map course times to rows and columns (example structure)
time_slots = {
    "8:00-9:00": 1,
    "09:00-10:00": 2,
    "10:00-11:00": 3,
    "11:00-12:00": 4,
    "12:00-13:00": 5,
    "13:00-14:00": 6,
    "14:00-15:00": 7,
    "15:00-16:00": 8,
    "16:00-17:00": 9,
}

days = ["一", "二", "三", "四", "五"]

# Display course data
for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, max_row=53, min_col=1, max_col=5, values_only=True), start=1):
    for col_idx, value in enumerate(row):
        label = tk.Label(content_frame, text=value if value is not None else "", borderwidth=1, relief="solid", padx=5, pady=5)
        label.grid(row=row_idx, column=col_idx, sticky="nsew", padx=2, pady=2)

    entry = tk.Entry(content_frame, width=15)
    entry.grid(row=row_idx, column=len(headers) - 1, padx=2, pady=2, sticky="nsew")

    def add_course(course_data=row, entry_widget=entry):
        student_id = entry_widget.get().strip()
        if not student_id:
            messagebox.showinfo("錯誤", "請輸入學號")
            return
        
        # Open personal schedule
        try:
            personal_schedule_wb = openpyxl.load_workbook(personal_schedule_path)
        except FileNotFoundError:
            personal_schedule_wb = openpyxl.Workbook()
            personal_schedule_wb.save(personal_schedule_path)
        
        # Create or get student's sheet
        if student_id not in personal_schedule_wb.sheetnames:
            student_sheet = personal_schedule_wb.create_sheet(student_id)
            student_sheet.append(["時間"] + days)  # Add header row
            for time in time_slots:
                student_sheet.append([time] + [""] * len(days))  # Initialize time slots
        else:
            student_sheet = personal_schedule_wb[student_id]

        # Extract course information
        course_name = course_data[0]
        course_time = course_data[2]  # Assuming "開課時間" is in the 3rd column
        
        # Find the row for the course time
        time_row = time_slots.get(course_time)
        if time_row is None:
            messagebox.showinfo("錯誤", "課程時間無效")
            return
        
        # Find the appropriate day column (you might need additional logic here)
        day_column = ...  # Logic to determine which day to place the course (not defined here)

        # Check if the slot is free before adding
        if student_sheet.cell(row=time_row + 1, column=day_column + 1).value:
            messagebox.showinfo("加選失敗", "該時間已有課程，無法加選")
        else:
            # Add course to the schedule
            student_sheet.cell(row=time_row + 1, column=day_column + 1, value=course_name)
            personal_schedule_wb.save(personal_schedule_path)
            messagebox.showinfo("加選成功", "課程已成功加入您的課表")

    confirm_button = tk.Button(content_frame, text="確認", command=add_course)
    confirm_button.grid(row=row_idx, column=len(headers), padx=2, pady=2, sticky="nsew")

Swindow.mainloop()