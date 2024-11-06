import tkinter as tk

def on_select_time():
    # 當按下按鈕時，顯示當前選擇的時間
    selected_time = time_var.get()
    current_time_label.config(text=f"當前選擇的時間：{selected_time}")

# 創建主視窗
root = tk.Tk()
root.title("時間選擇")

# 設定一個標籤顯示當前選擇的時間
label = tk.Label(root, text="請選擇上課時間")
label.pack(pady=10)

# 設定時間選項
time_options = ["08:00-10:00", "10:00-12:00", "12:00-14:00", "14:00-16:00", "16:00-18:00", "18:00-20:00"]

# 創建選擇時間的 OptionMenu
time_var = tk.StringVar(root)
time_var.set(time_options[0])  # 設定初始值

time_menu = tk.OptionMenu(root, time_var, *time_options)
time_menu.pack(pady=10)

# 顯示當前選擇的時間
current_time_label = tk.Label(root, text=f"當前時間：{time_var.get()}")
current_time_label.pack()

# 按鈕，當按下時更新顯示的時間
update_button = tk.Button(root, text="更新時間", command=on_select_time)
update_button.pack(pady=10)

# 啟動 Tkinter 主循環
root.mainloop()
