# 建立GUI 6.0
import tkinter as tk
import tkinter.font as tkFont
from tkinter import *
from tkinter import ttk
from tkinter import scrolledtext
from tkcalendar import * 
from datetime import datetime, timedelta
import os
import csv
import formulas
import pandas as pd

# import sys
# sys.path.append('D:\NTUmba\商管程式設計\期末專案')

###

file_path = "./schedule.csv" #存取檔案路徑
if os.path.isfile(file_path):
    os.remove(file_path)

class Meeting(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.grid()
        self.createWidgets()

    def createWidgets(self):
        f1 = tkFont.Font(size=14, family="Arial")
        
        # 建立主頁面
        self.labelIdentity = tk.Label(self, text='您的身份是:', height=1, width=10, font=f1, anchor='w')
        self.boxIdentity = ttk.Combobox(self, width=10, values=['會議發起者', '會議與會者'], font=f1)
        self.btnConfirm = tk.Button(self, text='確認', height=1, width=1, font=f1, command=self.clickBtnConfirm, activeforeground='#f00')
        self.btnResult = tk.Button(self, text='查看會議結果', height=1, width=1, font=f1, command=self.clickBtnResult, activeforeground='#f00')
        self.cvsMain = tk.Canvas(self, width = 600, height = 200, bg = "white")
                
        self.labelIdentity.grid(row=0, column=0, sticky = tk.NE + tk.SW)
        self.boxIdentity.grid(row=0, column=1, sticky = tk.NE + tk.SW)
        self.btnConfirm.grid(row=0, column=2, sticky = tk.NE + tk.SW)   
        self.btnResult.grid(row=1, column=0, columnspan=5, sticky = tk.NE + tk.SW)
        self.cvsMain.grid(row=10, column=0, columnspan=5, sticky = tk.NE + tk.SW)

    def clickBtnConfirm(self):
        self.valueIdentity = self.boxIdentity.get()
        # 會議發起者頁面
        if self.valueIdentity == '會議發起者':
            self.__init__
            self.createWidgets()
            
            self.lblIdentityHost = tk.Label(self, text = "您現在的身份是：會議發起者", height=2, width=1, font = ('Arial',16,'bold'))
            self.lblTitle = tk.Label(self, text = "會議名稱:", height=1, width=1, font = ('Arial',14), anchor='w')
            self.lblLength = tk.Label(self, text = "會議時長:", height=1, width=1,  font = ('Arial',14), anchor='w')
            self.lblLengthHour = tk.Label(self, text = "小時", height=1, width=1, font = ('Arial',14), anchor='w')
            self.lblLengthMin = tk.Label(self, text = "分鐘", height=1, width=1,  font = ('Arial',14), anchor='w')
            self.lblStartDate = tk.Label(self, text = "投票開始日期:", height=1, width=1, font = ('Arial',14), anchor='w')
            self.lblEndDate = tk.Label(self, text = "投票結束日期:", height=1, width=1,  font = ('Arial',14), anchor='w') 
            self.lblLocation = tk.Label(self, text = "會議地點:", height=1, width=1, font = ('Arial',14), anchor='w')
            self.lblStartTime = tk.Label(self, text = "投票開始時間:", height = 1, font = ('Arial',14), anchor='w')
            self.lblEndTime = tk.Label(self, text = "投票結束時間:", height = 1,  font = ('Arial',14), anchor='w') 
            self.txtTitle = tk.Text(self, height=1, width=20, font = ('Courier New',14))
            self.txtLength = tk.Text(self, height=1, width=20, font = ('Courier New',14))
            self.txtLocation = tk.Text(self, height=1, width=20, font = ('Courier New',14))
            self.calStart = DateEntry(self, year = 2023, month = 5, day = 19, setmode = 'day', date_pattern = 'yyyy/mm/dd', background = 'gray', foreground = 'white', borderwidth = 1, height=1)
            self.calEnd = DateEntry(self, year = 2023, month = 5, day = 19, setmode = 'day', date_pattern = 'yyyy/mm/dd', background = 'gray', foreground = 'white', borderwidth = 1, height=1)
            self.boxMeetHour = ttk.Combobox(self, width=10, values=['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24'])
            self.boxMeetMin = ttk.Combobox(self, width=10, values=['0', '30'])
            self.btnBuild = tk.Button(self, text="建立會議!", height=1, command=self.clickBtnBuild, font = ('Arial',14), activeforeground='#f00')
            hour_list = []
            for i in range(0,25):
                hour = str(i)+ ":00"
                hour_list.append(hour)
            self.boxHourStart = ttk.Combobox(self, width = 20, values=hour_list)
            self.boxHourEnd = ttk.Combobox(self, width = 20, values=hour_list)
            self.lblNth = tk.Label(self, text='', height=1, width=1)
       
            self.labelIdentity.grid(row=0, column=0, sticky = tk.NE + tk.SW)
            self.boxIdentity.grid(row=0, column=1, sticky = tk.NE + tk.SW)
            self.btnConfirm.grid(row=0, column=2, sticky = tk.NE + tk.SW)   
            self.btnResult.grid(row=1, column=0, columnspan=5, sticky = tk.NE + tk.SW)
            self.cvsMain.grid(row=10, column=0, columnspan=5, sticky = tk.NE + tk.SW)

            self.boxIdentity.current(0)
            
            self.lblIdentityHost.grid(row = 2, column = 0, columnspan = 5, sticky = tk.NE + tk.SW)
            
            self.lblTitle.grid(row = 3, column = 0, sticky = tk.NE + tk.SW)
            self.txtTitle.grid(row = 3, column = 1, columnspan = 4, sticky = tk.NE + tk.SW)
    
            self.lblLocation.grid(row = 4, column = 0, sticky = tk.NE + tk.SW)
            self.txtLocation.grid(row = 4, column = 1, columnspan = 4, sticky = tk.NE + tk.SW)
    
            self.lblLength.grid(row = 5, column = 0, sticky = tk.NE + tk.SW)
            self.boxMeetHour.grid(row = 5, column = 1, sticky = tk.NE + tk.SW)
            self.lblLengthHour.grid(row = 5, column = 2, sticky = tk.NE + tk.SW)
            self.boxMeetMin.grid(row = 5, column = 3, sticky = tk.NE + tk.SW)
            self.lblLengthMin.grid(row = 5, column = 4, sticky = tk.NE + tk.SW)
    
            self.lblStartDate.grid(row = 6, column = 0, sticky = tk.NE + tk.SW)
            self.calStart.grid(row = 6, column = 1, columnspan = 4, sticky = tk.NE + tk.SW)
    
            self.lblEndDate.grid(row = 7, column = 0, sticky = tk.NE + tk.SW)
            self.calEnd.grid(row = 7, column = 1,  columnspan = 4, sticky = tk.NE + tk.SW)
            
            self.lblStartTime.grid(row = 8, column = 0, sticky = tk.NE + tk.SW)
            self.boxHourStart.grid(row = 8, column = 1, sticky = tk.NE + tk.SW)
            self.lblNth.grid(row = 8, column = 2, sticky = tk.NE + tk.SW)
            self.lblEndTime.grid(row = 8, column = 3, sticky = tk.NE + tk.SW)
            self.boxHourEnd.grid(row = 8, column = 4, sticky = tk.NE + tk.SW)
            
            self.btnBuild.grid(row = 9, column = 0, columnspan = 5, sticky = tk.NE + tk.SW)
            
        # 會議與會者頁面
        elif self.valueIdentity == '會議與會者':
            fh = open('meeting.csv', "r") #存取檔案路徑
            csvFile = csv.DictReader(fh)
            for row in csvFile:
                self.m_startDate = row["start date"]
                self.m_endDate = row["end date"]
                self.m_startTime = row["start time"]
                self.m_endTime = row["end time"]

            date_delta = formulas.date_range(self.m_startDate, self.m_endDate)
            time_delta = formulas.get_time_range(self.m_startTime, self.m_endTime)
            current_date, current_time = date_delta[0], time_delta[0]
            date_values, time_values = [], []
            while current_date <= date_delta[1]:
                date = datetime.strftime(current_date, '%Y/%m/%d')
                date_values.append(date)
                current_date += timedelta(days=1)
            
            while current_time <= time_delta[1]:
                time = datetime.strftime(current_time, "%H:%M")
                time_values.append(time)
                current_time += timedelta(minutes=30)
            
            self.__init__
            self.createWidgets()
            
            self.lblIdentityAttendee = tk.Label(self, text = "您現在的身份是：會議與會者", height=2, width=1, font = ('Arial',16,'bold'))
            self.lblAttendant = tk.Label(self, text="您的姓名:", height=1, width=1, font=('Arial',14), anchor='w')
            self.lblEmail = tk.Label(self, text="您的電子信箱:", height=1, width=1, font=('Arial',14), anchor='w')
            self.txtAttendant = tk.Text(self, height = 1, width = 20, font=('Courier New',14))
            self.txtEmail = tk.Text(self, height = 1, width = 20, font=('Courier New',14))
            self.lblDate = tk.Label(self, text="選擇日期:", height=1, width=1, font=('Arial',14), anchor='w')
            self.lbltimeStart= tk.Label(self, text="選擇開始時間:", height=1, width=1, font=('Arial',14), anchor='w')
            self.lbltimeEnd= tk.Label(self, text="選擇結束時間:", height=1, width=1, font=('Arial',14), anchor='w')
            self.txttimeStart = tk.Text(self, height = 1, width = 20, font=('Courier New',14))
            self.txttimeEnd = tk.Text(self, height = 1, width = 20, font=('Courier New',14))
            self.btnTime = tk.Button(self, text="加入時間!", height=2, command=self.clickBtnJoin, font = ('Arial',14), activeforeground='#f00')

            self.boxChooseDate = tk.ttk.Combobox(self, width=10, values=date_values)
            self.boxChooseStartTime = tk.ttk.Combobox(self, width=10, values=time_values)
            self.boxChooseEndTime = tk.ttk.Combobox(self, width=10, values=time_values)
            self.boxIdentity.current(1)
            self.lblIdentityAttendee.grid(row = 2, column = 0, columnspan = 5, sticky = tk.NE + tk.SW)
            
            self.lblAttendant.grid(row=3, column=0, sticky = tk.NE + tk.SW)
            self.txtAttendant.grid(row=3,column=1, columnspan=4, sticky = tk.NE + tk.SW)
            
            self.lblEmail.grid(row=4, column=0, sticky = tk.NE + tk.SW)
            self.txtEmail.grid(row=4,column=1, columnspan=4, sticky = tk.NE + tk.SW)
            
            self.lblDate.grid(row=5,column=0, sticky = tk.NE + tk.SW)
            self.boxChooseDate.grid(row=5,column=1, columnspan = 4, sticky = tk.NE + tk.SW)
            
            self.lbltimeStart.grid(row=6,column=0, sticky = tk.NE + tk.SW)
            self.boxChooseStartTime.grid(row=6,column=1, columnspan = 4, sticky = tk.NE + tk.SW)
            
            self.lbltimeEnd.grid(row=7,column=0, sticky = tk.NE + tk.SW)
            self.boxChooseEndTime.grid(row=7,column=1, columnspan = 4, sticky = tk.NE + tk.SW)

            self.btnTime.grid(row=8,column=0, rowspan=2, columnspan=5, sticky = tk.NE + tk.SW)
            
    # 建立會議
    def clickBtnBuild(self):
        
        self.cvsMain.delete('all')
        text_list= [self.txtTitle, self.txtLocation]
        entry_list = [self.calStart, self.calEnd, self.boxMeetHour, self.boxMeetMin, self.boxHourStart, self.boxHourEnd]
        # 偵測空格+偵測開始時間晚於結束時間
        if check_empty_textbox(text_list) is False or check_empty_combobox(entry_list) is False:
            self.cvsMain.create_text(330,30,text="表格尚未填寫完成",font = ('Arial',20, 'bold'), fill='red')
        elif check_date_range(self.calStart.get(), self.calEnd.get()) is False or check_time_range(self.boxHourStart.get(), self.boxHourEnd.get()) is False:
            if check_date_range(self.calStart.get(), self.calEnd.get()) is False and check_time_range(self.boxHourStart.get(), self.boxHourEnd.get()) is False:
                self.cvsMain.create_text(330,30,text="日期填寫有誤、時間填寫有誤",font = ('Arial',20, 'bold'), fill='red')
            elif check_time_range(self.boxHourStart.get(), self.boxHourEnd.get()) is True:
                self.cvsMain.create_text(330,30,text="日期填寫有誤",font = ('Arial',20, 'bold'), fill='red')
            else:
                self.cvsMain.create_text(330,30,text="時間填寫有誤",font = ('Arial',20, 'bold'), fill='red')
        else:
            self.m_title = self.txtTitle.get("0.0", tk.END)
            self.m_location = self.txtLocation.get("0.0", tk.END)
            self.m_startDate = self.calStart.get()
            self.m_endDate = self.calEnd.get()
            self.int_hour = self.boxMeetHour.get()
            self.int_min = self.boxMeetMin.get()
            self.m_hour = self.int_hour + '小時'
            if self.int_min == '30':
                self.m_min = self.int_min + '分鐘'
            else:
                self.m_min = ''
            self.m_startTime = self.boxHourStart.get()
            self.m_endTime = self.boxHourEnd.get()
            
            with open('meeting.csv', "w", newline='') as test: #存取檔案路徑
                writer = csv.writer(test)
                # 輸入標題
                csv_head = ['title', 'location', 'start date', 'end date', 'start time', 'end time', 'hour', 'min']
                writer.writerow(csv_head)
                writer.writerow([self.m_title, self.m_location, self.m_startDate, self.m_endDate, self.m_startTime, self.m_endTime, self.int_hour, self.int_min])

            self.output = [self.m_title, self.m_hour, self.m_min, self.m_location, self.m_startDate, self.m_endDate, self.m_startTime, self.m_endTime, self.int_hour, self.int_min]
            return self.output

    #查看會議結果頁面
    def clickBtnResult(self):
        fh = open('meeting.csv', "r") #存取檔案路徑
        csvFile = csv.DictReader(fh)
        for row in csvFile:
            self.m_title = row["title"]
            self.m_location = row["location"]
            self.m_startDate = row["start date"]
            self.m_endDate = row["end date"]
            self.m_startTime = row["start time"]
            self.m_endTime = row["end time"]
            self.int_hour = row["hour"]
            self.int_min = row["min"]
        date_delta = formulas.get_date_range(self.m_startDate, self.m_endDate)
        date_schedule = []
        for i in date_delta:
            date_schedule.append(formulas.Schedule(i).schedule)
        meeting_length = int(self.int_hour) * 60 + int(self.int_min)
        self.data = pd.read_csv('schedule.csv') #存取檔案路徑
        name_num = self.data['name'].nunique()
        # 讀取每一個人的available time
        fh = open('schedule.csv', "r") #存取檔案路徑
        test = csv.DictReader(fh)
        # 把每一行的時段加進對應天數的清單中
        for row in test:
            date_schedule[date_delta.index(row['date'])] += \
                row['time string'].split(',')

        # 接著要找所有人都可以出席的時間
        # Step 1: 先把對應天數的清單內的所有時段都轉成浮點數，接著排序
        date_schedule = [[float(element) for element in sublist] for sublist in date_schedule]
        for i in date_schedule:
            i.sort()
        # Step 2: 建立一個二維空清單，若投票區間有n天，裡面就有n個子清單
        overlap_schedule = [[] for _ in range(len(date_schedule))]
        # Step 3: 把所有人都可以的時間（有m個人，就找出現m次的時間點）加進對應的子清單
        for i in range(len(date_schedule)):
            for j in date_schedule[i]:
                if j not in overlap_schedule[i] and \
                        date_schedule[i].count(j) == int(name_num):
                    
                    overlap_schedule[i].append(j)
        # Step 4: 尋找每個子清單中符合會議時長（meeting_length）的時間區段
        meeting_output = []
        for k in range(len(overlap_schedule)):
            result = []
            time_len = int(meeting_length) // 30 + 1  # 後面+1因為n個半小時需要n+1個時間點
            for i in range(len(overlap_schedule[k]) - time_len + 1):
                subset = overlap_schedule[k][i:i + time_len]
                if all(subset[j] + 0.5 == subset[j + 1] for j in range(time_len - 1)):
                    result.append(subset)
            if len(result) == 0:
                continue
            else:
                for period in result:
                    beg = formulas.float2time(period[0])
                    end = formulas.float2time(period[-1])
                    meeting_output.append([f'{date_delta[k]} {beg}', f"{date_delta[k]} {end}"])

        self.destroy()
        self.__init__()
        self.createWidgets()
        meetingOutput = []
        for ln in range(len(meeting_output)):
            meetingOutput.append(meeting_output[ln][0] + ' - ' + meeting_output[ln][1])
        
        if meetingOutput != []:
            self.btnGoogle = tk.Button(self, text="建立Google Calendar!", height=1, command=self.clickBtnCal, font = ('Arial',14), activeforeground='#f00')
            self.lblChoose = tk.Label(self, text="選擇會議時間:", height=1, width=1, font=('Arial',14), anchor='w')
            self.boxTime = ttk.Combobox(self, width=10, values=meetingOutput)
            self.lblChoose.grid(row=2, column=0, sticky = tk.NE + tk.SW)
            self.boxTime.grid(row=2, column=1, columnspan=2, sticky = tk.NE + tk.SW)
            self.btnGoogle.grid(row=2, column=3, columnspan=2, sticky = tk.NE + tk.SW)
        else:
            self.lblNoTime = tk.Label(self, text="沒有所有人皆能出席的時間！", height=2, width=1, font=('Arial',16, 'bold'), fg='#f00')
            self.lblNoTime.grid(row=2, column=0, columnspan=5, sticky = tk.NE + tk.SW)
        
        self.cvsMain.delete('all')
        self.cvsMain.create_text(15,30,text="會議名稱: "+ self.m_title, anchor='w', font = ('Arial',16))
        if self.int_min == '30':
            min = self.int_min + '分鐘'
        else:
            min = ''
        self.cvsMain.create_text(15,46,text="會議時長: "+ self.int_hour + "小時" + min, anchor='w', font = ('Arial',16))
        self.cvsMain.create_text(15,80,text="會議地點: "+ self.m_location, anchor='w', font = ('Arial',16))
        # self.cvsMain.create_text(15,95,text="投票開始日期: "+output[4], anchor='w', font = ('Arial',16))
        # self.cvsMain.create_text(15,119,text="投票結束日期: "+output[5], anchor='w', font = ('Arial',16)) 
        # self.cvsMain.create_text(15,143,text="投票開始時間: "+output[6], anchor='w', font = ('Arial',16))
        # self.cvsMain.create_text(15,167,text="投票結束時間: "+output[7], anchor='w', font = ('Arial',16)) 
        # if not meeting_output:
        #     self.cvsMain.create_text(15,95,text="沒有所有人皆能出席的時間！", anchor='w', font = ('Arial',16))
        # else:
        #     self.cvsMain.create_text(15,95,text="以下為所有人皆能出席的時間，請選擇最終決定開會時段：")
        #     var = StringVar()
        #     var.set(output[0])
        #     for item in meeting_output:
        #         button = Radiobutton(self.cvsMain, text=item, variable=var, value=item)
        #         button.pack(anchor=W)

    # 與會者加入會議
    def clickBtnJoin(self):
        self.name = self.txtAttendant.get("0.0", tk.END)
        self.email = self.txtEmail.get("0.0", tk.END)
        self.date = self.boxChooseDate.get()
        self.startTime = self.boxChooseStartTime.get()
        self.endTime = self.boxChooseEndTime.get()
        file_path = "./schedule.csv" #存取檔案路徑
        if os.path.isfile(file_path):
            with open(file_path, "a+", newline='') as test:
                writer = csv.writer(test)
                writer.writerow([self.name, self.date, self.startTime, self.endTime, formulas.time_str(self.startTime, self.endTime)])
        else:
            with open('schedule.csv', "w", newline='') as test: #存取檔案路徑
                writer = csv.writer(test)
                # 輸入標題
                csv_head = ['name', 'date', 'start time', 'end time', 'time string']
                writer.writerow(csv_head)
                writer.writerow([self.name, self.date, self.startTime, self.endTime, formulas.time_str(self.startTime, self.endTime)])

    def clickBtnCal(self):
        self.final_decision = self.boxTime.get()
        self.final_beg, self.final_end = self.final_decision.split(' - ')
        api_begin = formulas.str2iso(self.final_beg)
        api_end = formulas.str2iso(self.final_end)
        print(api_begin, api_end)

def check_empty_textbox(list):
    for entry in list:
        if len(entry.get("0.0", tk.END))==0:
            return False
    return True

def check_empty_combobox(list):
    for entry in list:
        if len(entry.get()) == 0:
            return False
    return True

def check_date_range(start_date, end_date):
    start = datetime.strptime(start_date, '%Y/%m/%d')
    end = datetime.strptime(end_date, '%Y/%m/%d')
    if start > end:
        return False
    return True

def check_time_range(start_time, end_time):
    start = datetime.strptime(start_time, '%H:%M')
    end = datetime.strptime(end_time, '%H:%M')
    if start > end:
        return False
    return True

m = Meeting()
m.master.title("Meeting Arrangement")
m.mainloop()
