from tkinter import *
import tkinter as tk
import openpyxl
from tkinter import messagebox

Swindow = Tk()
Swindow.title("選課系統")
Swindow.geometry("1500x1000")

# Canvas and Scrollbar setup
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

def on_mouse_wheel(event):
    if event.delta > 0:
        canvas.yview_scroll(-1, "units")
    elif event.delta < 0:
        canvas.yview_scroll(1, "units")

canvas.bind_all("<MouseWheel>", on_mouse_wheel)
canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

def adjust_column_widths(event=None):
    canvas_width = canvas.winfo_width()
    col_width = canvas_width // len(headers)
    for col, header in enumerate(headers):
        label = tk.Label(content_frame, text=header, font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=col_width//10)
        label.grid(row=0, column=col, sticky="nsew", padx=2, pady=2)

canvas.bind("<Configure>", adjust_column_widths)

# Button to open a new window
def open_new_window():
    new_window = Toplevel(Swindow)
    new_window.title("課表頁面")
    new_window.geometry("400x300")
    label = Label(new_window, text="個人課表")
    label.pack(pady=20)
    close_button = Button(new_window, text="關閉", command=new_window.destroy)
    close_button.pack(pady=10)

button_pop = Button(Swindow, text="課表頁面", command=open_new_window)
button_pop.place(x=10, y=10)

# Load the Excel file and worksheets
path = '/Users/hayashitogi/Documents/GitHub/group_7/資料庫.xlsx'
workbook = openpyxl.load_workbook(path)
worksheet_courses = workbook["課程"]
worksheet_students = workbook["學生"]

# Headers
headers = ["課程名稱", "課程代碼", "開課時間", "上課地點", "授課教授", "加退選匡"]

# Display course data
for row_idx, row in enumerate(worksheet_courses.iter_rows(min_row=2, max_row=53, min_col=1, max_col=5, values_only=True), start=1):
    for col_idx, value in enumerate(row):
        label = tk.Label(content_frame, text=value if value is not None else "", borderwidth=1, relief="solid", padx=5, pady=5)
        label.grid(row=row_idx, column=col_idx, sticky="nsew", padx=2, pady=2)
    
    # Entry and confirmation button for each course row
    entry = tk.Entry(content_frame, width=15)
    entry.grid(row=row_idx, column=len(headers) - 1, padx=2, pady=2, sticky="nsew")

    def number_search(entry=entry):
        student_number = entry.get()
        # Verify if the student exists and retrieve their course schedule
        student_found = False
        for row in worksheet_students.iter_rows(min_row=2, values_only=True):
            name, id_, schedule_path = row[:3]  # 解包每一行的資料
            messagebox.showinfo("成功", id_)
            if id_ == entry:
                student_found = True
                student_schedule = schedule_path  # 3rd column is schedule
                messagebox.showinfo("成功", "已找到該學生")
                #print(f"Schedule for {student_number}: {student_schedule}")
                break
            else: 
                messagebox.showinfo("錯誤", "未找到該學生")
                break

        if student_found:
            entry.delete(0, END)  # Clear the entry after confirming
            # Further code can display the schedule if needed

    confirm_button = tk.Button(content_frame, text="確認", command=number_search)
    confirm_button.grid(row=row_idx, column=len(headers), padx=2, pady=2, sticky="nsew")

Swindow.mainloop()