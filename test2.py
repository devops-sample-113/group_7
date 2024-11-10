import tkinter as tk

root = tk.Tk()
root.geometry("500x500")  # 設定初始視窗大小

# 建立一個可以滾動的框架
canvas = tk.Canvas(root)
scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
frame = tk.Frame(canvas)

# 在畫布上建立滾動區域
canvas.create_window((0, 0), window=frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

# 放一些內容以便測試滾動
for i in range(50):
    tk.Label(frame, text=f"Label {i}").pack()

# 設定滾動條
scrollbar.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)

frame.update_idletasks()

# 更新畫布的區域大小
canvas.config(scrollregion=canvas.bbox("all"))

root.mainloop()