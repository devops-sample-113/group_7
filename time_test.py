def extract_keywords(course_time):
    # 提取星期幾（例如：星期二）和時間區間（例如：13:00-15:00）
    day, time_range = course_time.split(" ")

    # 提取星期幾的關鍵字（例如 "星期二" 變成 "二"）
    day_keyword = day[2]  # 這裡取的是「二」字，假設格式固定為「星期X」

    # 提取開始和結束時間（例如 "13:00-15:00"）
    start_time, end_time = time_range.split('-')

    # 提取並將開始時間和結束時間轉換為整數
    start_hour = int(start_time.split(":")[0])  # 提取 "13" 並轉為整數
    end_hour = int(end_time.split(":")[0])      # 提取 "15" 並轉為整數

    # 返回關鍵字詞及整數時間
    return day_keyword, start_hour, end_hour

def take_time(course_time):
    # 提取星期幾（例如：星期二）和時間區間（例如：13:00-15:00）
    _, time_range = course_time.split(" ")

    # 提取開始和結束時間（例如 "13:00-15:00"）
    start_time, end_time = time_range.split('-')

    # 提取並將開始時間和結束時間轉換為整數
    start_hour = int(start_time.split(":")[0])  # 提取 "13" 並轉為整數
    end_hour = int(end_time.split(":")[0])      # 提取 "15" 並轉為整數

    # 返回關鍵字詞及整數時間
    return start_hour, end_hour

# 測試例子
course_time = "星期三 08:00-10:00"
day_keyword, start_hour, end_hour = extract_keywords(course_time)

print(f"星期幾: {day_keyword}, 開始小時: {start_hour}, 結束小時: {end_hour}")
