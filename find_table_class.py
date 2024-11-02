import openpyxl

# 設定 Excel 檔案路徑
file_path = "C:\\Users\\User\\OneDrive - 逢甲大學\\文件\\GitHub\\group_7\\個人課表\\D0987654.xlsx"
course_sheet = openpyxl.load_workbook(file_path)["課表"]

def find_course_in_schedule(sheet, course_name):
    # 假設課表是從 A1 開始的 5 行 8 列範圍
    for row in sheet.iter_rows(min_row=1, max_row=9, min_col=1, max_col=9, values_only=True):
        for cell in row:
            if cell == course_name:
                return True  # 找到課程，返回 True
    return False  # 如果遍歷完成仍未找到，返回 False

# 示例：查找課程名稱
course_name_to_find = "軟體測試"  # 您要查找的課程名稱

if find_course_in_schedule(course_sheet, course_name_to_find):
    print(f"課程 {course_name_to_find} find")
else:
    print(f"課程 {course_name_to_find} not find")
