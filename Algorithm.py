import csv
from datetime import datetime, timedelta


class Attendant:
    # 會議參加者的名字（name）和電子信箱（email）
    def __init__(self, name, email):
        self.name = name
        self.email = email
        self.schedule = []


class Meeting:
    # 會議本身的標題（time）、時長（length）、地點（location）
    def __init__(self, title, length, location):
        self.title = title
        self.length = length
        self.location = location


class Schedule:
    def __init__(self, date):
        self.date = date
        self.schedule = []


# 創建一個字典儲存所有會議參加者的姓名和email
def get_attendant_info():
    full_ppl_info = dict()
    print("請輸入參加者姓名與email，輸入None則結束")
    while True:  # 輸入None即跳出
        att_n = input("參加者姓名：")  # 未來要做除錯
        if att_n == "None":
            break
        else:
            e = input("參加者的電子信箱：")  # 未來要做除錯
            a = Attendant(att_n, e)
            # 將參加者的資訊做成Attendant資料並加入字典，key為其名字
            full_ppl_info[att_n] = a
    return full_ppl_info


# 創建可投票日期的清單
# 輸入 2023/05/17, 2023/05/20 回傳 [5/17, 5/18, 5/19, 5/20]
def get_date_range(start_date, end_date):  # 還要寫除錯
    date_format = "%Y/%m/%d"  # 輸入格式2023/5/17
    output_format = '%m/%d'
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


# 將浮點數轉為AM/PM格式輸出時間
# e.g. 輸入 8.5 回傳 AM 8:30；輸入 22.0 回傳 PM 10:00
def float2time(float_num):
    hour = int(float_num)
    minute = int((float_num - hour) * 60)
    if hour < 12:
        p = "AM"
    else:
        p = "PM"
        if hour > 12:
            hour -= 12
    return f"{p} {hour}:{minute:02d}"


# 輸入會議的基本資訊
m_title = input("請輸入會議名稱:")
m_length = input("請輸入會議時長（分鐘，以30為一單位）:")  # 格式:MM
m_location = input("請輸入會議地點:")
meeting = Meeting(m_title, m_length, m_location)

# 輸入投票開會的日期範圍
s_date = input('請輸入投票區間開始日期（yyyy/mm/dd）：')  # 格式：2023/5/17
e_date = input('請輸入投票區間結束日期（yyyy/mm/dd）：')

# 建立名字的清單
f = get_attendant_info()
name_list = list(f.keys())

# 叫出所有可投票的日期清單
date_range = get_date_range(s_date, e_date)
date_schedule = []
for i in date_range:
    date_schedule.append(Schedule(i).schedule)

# 建立空白的csv檔
with open('schedule.csv', "w", newline='') as test:
    writer = csv.writer(test)
    # 輸入標題
    csv_head = ['name', 'date', 'start time', 'end time', 'time string']
    writer.writerow(csv_head)

    # 輸入每個人的available time
    print('::請輸入每一位與會者在每一天分別可以的開始時間')
    print('::若有多段空閒時間可多次輸入')
    print('::欲輸入下一天或下一位與會者的有空時間，請輸入n')
    for n in name_list:
        for d in date_range:
            while True:
                # 輸入格式：8:00
                start = input(f'{n}在{d}有空的開始時間：')
                if start == 'n':
                    break
                # 輸入格式：8:00
                end = input(f'{n}在{d}有空的結束時間：')
                writer.writerow([n, d, start, end, time_str(start, end)])

# 讀取每一個人的available time
fh = open('schedule.csv', "r")
test = csv.DictReader(fh)
# 把每一行的時段加進對應天數的清單中
for row in test:
    date_schedule[date_range.index(row['date'])] += \
        row['time string'].split(',')

# 接著要找所有人都可以出席的時間
# Step 1: 先把對應天數的清單內的所有時段都轉成浮點數，接著排序
date_schedule = [[float(element) for element in sublist] for
                 sublist in date_schedule]
for i in date_schedule:
    i.sort()
# Step 2: 建立一個二維空清單，若投票區間有n天，裡面就有n個子清單
overlap_schedule = [[] for _ in range(len(date_schedule))]
# Step 3: 把所有人都可以的時間（有m個人，就找出現m次的時間點）加進對應的子清單
for i in range(len(date_schedule)):
    for j in date_schedule[i]:
        if j not in overlap_schedule[i] and \
                date_schedule[i].count(j) == len(name_list):
            overlap_schedule[i].append(j)
# Step 4: 尋找每個子清單中符合會議時長（meeting.length）的時間區段
output = []
for k in range(len(overlap_schedule)):
    result = []
    time_len = int(meeting.length) // 30 + 1  # 後面+1因為n個半小時需要n+1個時間點
    for i in range(len(overlap_schedule[k]) - time_len + 1):
        subset = overlap_schedule[k][i:i + time_len]
        if all(subset[j] + 0.5 == subset[j + 1] for j in range(time_len - 1)):
            result.append(subset)
    if len(result) == 0:
        continue
    else:
        for period in result:
            beg = float2time(period[0])
            end = float2time(period[-1])
            output.append(f'{date_range[k]} {beg} - {end}')

# 最終輸出
print(" ")
print("以下是您的會議資訊：")
print(f"會議名稱：{meeting.title}")
print(f"會議時長：{meeting.length}分鐘")
print(f"會議地點：{meeting.location}")
print(" ")
print("以下是參加者的資訊")
for k in f:
    print(f"姓名：{k} Email：{f[k].email}")
print(" ")
if not output:
    print("沒有所有人皆能出席的時間！")
else:
    print("以下為所有人皆能出席的時間：")
    for line in output:
        print(line)
        
