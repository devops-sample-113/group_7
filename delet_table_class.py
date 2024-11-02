import openpyxl

# 設定 Excel 檔案路徑
file_path = "C:\\Users\\User\\OneDrive - 逢甲大學\\文件\\GitHub\\group_7\\個人課表\\D0987654.xlsx"
course_sheet = openpyxl.load_workbook(file_path)["課表"]

def remove_all_courses(sheet, course_name):
    found = False  # 用於追蹤是否有找到該課程
    # 假設課表是從 A1 開始的 9 行 9 列範圍
    for row in sheet.iter_rows(min_row=1, max_row=9, min_col=1, max_col=9):
        for cell in row:
            if cell.value == course_name:
                cell.value = None  # 將儲存格的值設為 None（移除內容）
                found = True  # 標記已找到課程
    return found  # 返回是否有移除任何內容

# 示例：移除所有出現的課程名稱
course_name_to_remove = "軟體測試"  # 您要移除的課程名稱

if remove_all_courses(course_sheet, course_name_to_remove):
    print(f"課程 {course_name_to_remove} 已移除")
else:
    print(f"課程 {course_name_to_remove} 未找到")

# 儲存變更到 Excel 檔案
course_sheet.parent.save(file_path)
