from tkinter import *
import openpyxl

top = Tk()
top.title("選課系統")
top.geometry("1500x1000")  # 設置主視窗的大小

def fetch_data_from_excel():
    wb = openpyxl.load_workbook('/Users/hayashitogi/Documents/GitHub/group_7/資料庫.xlsx', data_only=True)
    s1 = wb['課程']
    data = ""
    for row in s1.iter_rows(min_row=1, min_col=1, max_col=4, max_row=5):  # 假設只取 4 列
        row_data = [cell.value for cell in row]
        data += "\t".join(str(value) for value in row_data) + "\n"  # 將每個欄位的值用 Tab 隔開
    wb.close()
    return data

def update_text_widget():
    data_text.config(state=NORMAL)  # 設置為可編輯，方便插入新資料
    data_text.delete(1.0, END)      # 清除舊資料
    data_text.insert(END, fetch_data_from_excel())  # 插入 Excel 資料
    data_text.config(state=DISABLED)  # 設置為不可編輯

def open_new_window():
    # 創建新的 Toplevel 視窗
    new_window = Toplevel(top)
    new_window.title("課表頁面")
    new_window.geometry("400x300")  # 設置新視窗的大小

    # 在新視窗上添加一些元件
    label = Label(new_window, text="個人課表")
    label.pack(pady=20)

    close_button = Button(new_window, text="關閉", command=new_window.destroy)
    close_button.pack(pady=10)

# 顯示 Excel 資料的 Text 小部件
data_text = Text(top, wrap=WORD, width=800, height=300)
data_text.place(x=10, y=80)
data_text.config(state=DISABLED)  # 初始設置為不可編輯

# 按鍵，點擊後開啟新頁面
button_pop = Button(top, text="課表頁面", command=open_new_window)
button_pop.place(x=10, y=10)

top.mainloop()
