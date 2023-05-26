# Formulas 2.0

import csv
from datetime import datetime, timedelta

# Function 1：host 發起會議
# 輸入資訊：time, length, location
# 串接 Google Sheet，將輸入會議資訊寫入試算表
# 創立新的 meeting A 試算表，產生 Key 編號

# Function 2: attendee 填入時間
# 輸入 Key 編號
# 透過 Google Sheet API，抓出 meeting A 試算表
# 輸入開會有空時間: name, available time
# 將輸入開會有空時間寫入 meeting A 試算表

# Function 3: 查看最終開會時間
# 輸入 Key 編號
# 串接 Google Sheet API，將對應的 meeting A 試算表抓下來，計算最終可以的開會時間
# 輸出所有可以的開會時間
# 輸入最終想要的開會時間編號
#  meeting A 試算表抓出與會者姓名與信箱，串接 Google Calendar API，自動寄送會議邀請

class Schedule:
    def __init__(self, date):
        self.date = date
        self.schedule = []

def date_range(start_date, end_date):
    date_format = "%Y/%m/%d"  # 輸入格式2023/05/17
    d_range = []

    # 將輸入的開始與結束時間轉換為datetime資料
    start_datetime = datetime.strptime(start_date, date_format)
    end_datetime = datetime.strptime(end_date, date_format)

    # 將開始和結束日期加入到清單中
    d_range.append(start_datetime)
    d_range.append(end_datetime)

    return d_range

def get_date_range(start_date, end_date):  # 還要寫除錯
    date_format = "%Y/%m/%d"  # 輸入格式2023/5/17
    output_format = '%Y/%m/%d'
    d_range = []

    # 將輸入的開始與結束時間轉換為datetime資料
    start_datetime = datetime.strptime(start_date, date_format)
    end_datetime = datetime.strptime(end_date, date_format)

    # 將開始日期加入到清單中
    d_range.append(start_datetime.strftime(output_format))

    # 利用迴圈將每一天加入清單中
    current_date = start_datetime
    while current_date < end_datetime:
        current_date += timedelta(days=1)
        d_range.append(current_date.strftime(output_format))

    return d_range


def get_time_range(start_time, end_time):
    time_format = "%H:%M"
    t_range = []

    start_time = datetime.strptime(start_time, time_format)
    end_time = datetime.strptime(end_time, time_format)

    t_range.append(start_time)
    t_range.append(end_time)

    return t_range

# 輸入時間字串回傳浮點數
# e.g. 輸入 8:00，回傳 8.0；輸入10:30，回傳 10.5
def time2float(t):
    h, m = int(t.split(':')[0]), int(t.split(':')[1])
    if m == 30:
        s = h + 0.5
    else:
        s = h
    return s

# 輸入兩個時間點，輸出兩者之間每半小時的清單
# e.g. 輸入 8:00, 10:30，回傳 [8.0, 8.5, 9.0, 9.5, 10.0, 10.5]
def time_str(t1, t2):
    t_list = []
    s1, s2 = time2float(t1), time2float(t2)
    t_list.append(s1)
    current_time = s1
    while current_time < s2:
        current_time += 0.5
        t_list.append(current_time)
    t_str = ",".join(map(str, t_list))
    return t_str

def float2time(float_num):
    hour = int(float_num)
    minute = int((float_num - hour) * 60)
    return f"{hour}:{minute:02d}"

# 將日期和時間轉換並輸出成iso format供google calendar API使用
# date is a "yyyy/mm/dd hh:mm" string
# 輸入日期如「2023/5/23 8:30」會輸出「2023-05-23T08:30:00」
def str2iso(date):
    time_dt = datetime.strptime(date, "%Y/%m/%d %H:%M")
    time_iso = time_dt.isoformat(timespec="seconds")
    return time_iso
