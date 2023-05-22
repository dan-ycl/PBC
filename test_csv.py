import csv

name_list = ['A', 'B']
date_range = ['5/17', '5/18']

def time2float(t):
    h, m = int(t.split(':')[0]), int(t.split(':')[1])
    if m == 30:
        s = h + 0.5
    else:
        s = h
    return s

def time_str(t1, t2):
    t_list= []
    s1, s2 = time2float(t1), time2float(t2)
    t_list.append(s1)
    current_time = s1
    while current_time < s2:
        current_time += 0.5
        t_list.append(current_time)
    t_str = ",".join(map(str, t_list))
    return t_str


# 建立空白的csv檔
with open('schedule.csv', "w", newline='' ) as test:
    writer = csv.writer(test)
    # 輸入標題
    csv_head = ['name', 'date', 'start time', 'end time', 'time string']
    writer.writerow(csv_head)

# 輸入每個人的available time
    print('請輸入每一位與會者在每一天分別可以的開始時間')
    print('若有多段空閒時間可多次輸入')
    print('欲輸入下一天或下一位與會者的有空時間，請輸入N')
    for name in name_list:
        for date in date_range:
            while True:
                # 輸入格式：8:00
                start = input(f'{name}在{date}有空的時間')
                if start == 'n':
                    break
                # 輸入格式：8:00
                end = input(f'{name}在{date}有空的結束時間')
                writer.writerow([name, date, start, end, time_str(start, end)])


# 讀取每一個人的available time
fh = open('schedule.csv', "r")
test = csv.DictReader(fh)
all_list = []
result = []
for row in test:
    all_list.append(row['time string'].split(',')) 
for i in all_list:
    for j in i:
        result.append(float(j))


print(result)


