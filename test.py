import pandas as pd
import tkinter as tk
from tkinter import ttk

# 讀取 Excel 文件
df = pd.read_excel('data.xlsx')  # 假設文件名為 'data.xlsx'

# 將數據進行一些操作，比如新增一列
df['New_Column'] = df['Existing_Column'] * 2  # 假設有一列叫 'Existing_Column'

# 寫回 Excel 文件
df.to_excel('modified_data.xlsx', index=False)  # 保存為 'modified_data.xlsx'

# 使用 tkinter 顯示數據
def display_data(dataframe):
    # 創建主窗口
    root = tk.Tk()
    root.title("Excel Data Display")

    # 創建一個 Treeview 來顯示表格數據
    tree = ttk.Treeview(root)
    tree["columns"] = list(dataframe.columns)
    tree["show"] = "headings"

    # 設置表頭
    for column in tree["columns"]:
        tree.heading(column, text=column)

    # 插入數據
    for _, row in dataframe.iterrows():
        tree.insert("", "end", values=list(row))

    # 設置滾動條
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    tree.pack(fill="both", expand=True)

    # 啟動主循環
    root.mainloop()

# 顯示數據
display_data(df)
