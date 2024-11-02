workbook=openpyxl(path)
worksheet=workbook.active

    ##個人課表顯示
# 讀取 1-10 行，A-F 列的內容並顯示在 Label 中
for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=10, min_col=1, max_col=6, values_only=True)):
    for col_idx, value in enumerate(row):
        label = tk.Label(Swindow, text=value if value is not None else "", borderwidth=1, relief="solid", padx=5, pady=5)
        label.grid(row=row_idx, column=col_idx, sticky="nsew", padx=2, pady=2)

# 自動調整列寬
for col in range(6):
    Swindow.grid_columnconfigure(col, weight=1)