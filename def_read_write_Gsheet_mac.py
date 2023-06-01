# 若下載後於個人電腦，請自行修改路徑如： credentials.json、meeting.csv、schedule.csv (ex. C:\\your\\own\\path\\file.csv)
import google.auth
from google.oauth2.service_account import Credentials
import pandas as pd
import pygsheets
import gspread
import os

### 函數區域
# 查詢 meeting.csv 的資料，根據會議名稱回傳開始、結束日期
# 主程式呼叫 1 次
def sendback_date(meeting_name):
    gc = pygsheets.authorize(service_account_file = cred_path)
    sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/18THYdAFqANCRORZkBp7tDgGgjXGeIWdLFsJK_0ez0qM/edit#gid=0')
    worksheet = sheet.worksheet('title', 'Overview')
    start_date = None
    end_date = None
    start_time = None
    end_time = None
    
    # 獲取 A 欄的數據
    column_data = worksheet.get_col(1)
    row_index = None

    for index, value in enumerate(column_data):
        if value == meeting_name:
            row_index = index + 1
            break
    if row_index:
        start_date = worksheet.get_value('C'+str(row_index))
        end_date = worksheet.get_value('D'+str(row_index))
        start_time = worksheet.get_value('E'+str(row_index))
        end_time = worksheet.get_value('F'+str(row_index))
        
    # else:
        # print('找不到會議名稱')
    
    return start_date, end_date, start_time, end_time
    
    



# Function 2-1: 將 schedule.csv 資料延續貼到 Gsheet
# 主程式沒有呼叫
def write_csv_to_GSheet(meeting_name, cred_path, input_csv):
    # 利用 pygsheets 開啟 GoogleSheet
    gc = pygsheets.authorize(service_account_file = cred_path)
    sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1h9r2dgUJD5Fze36lkCy0sbpeQAw4V0PJlcc0cVLU4iw/edit#gid=0')
    worksheet = sheet.worksheet('title', meeting_name)

    # sample CSV file, use your path
    df = pd.read_csv(input_csv, encoding = 'utf_8_sig')  # utf_8_sig  在用import codecs 前，big5 OK!
    cells = worksheet.get_all_values(include_tailing_empty_rows=False, include_tailing_empty=False, returnas='matrix')
    next_row = len(cells)+1
    worksheet.set_dataframe(df, start = (next_row, 1), copy_head = False)

    return
    

# Function 2-2: 將 meeting.csv 資料延續貼到 Gsheet，並在 schedule 的 spreadsheet 根據會議名稱建立新的分頁
# 主程式呼叫 1 次
def write_meeting_csv_to_GSheet(sheet_name, cred_path, input_meeting):
    # 利用 pygsheets 開啟 GoogleSheet
    gc = pygsheets.authorize(service_account_file = cred_path)
    sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/18THYdAFqANCRORZkBp7tDgGgjXGeIWdLFsJK_0ez0qM/edit#gid=0')  # meeting.csv
    worksheet = sheet.worksheet('title', sheet_name)

    df = pd.read_csv(input_meeting, encoding = 'utf_8_sig')  # utf_8_sig  
    cells = worksheet.get_all_values(include_tailing_empty_rows=False, include_tailing_empty=False, returnas='matrix')
    last_row = str(len(cells)+1)
    worksheet.set_dataframe(df, start = "A" + last_row ,copy_head=False)

    ### 建立雲端 schedule 的新分頁
    sheet_schedule = gc.open_by_url('https://docs.google.com/spreadsheets/d/1h9r2dgUJD5Fze36lkCy0sbpeQAw4V0PJlcc0cVLU4iw/edit#gid=0')  # schedule.csv
    worksheet_schedule = sheet_schedule.add_worksheet(df.iloc[0,0])  # 填入會議名稱
    worksheet_schedule.resize(500, 7)  # 預設工作表尺寸500列 7欄
    worksheet_schedule.update_value('A1', 'name')
    worksheet_schedule.update_value('B1', 'email')
    worksheet_schedule.update_value('C1', 'date')
    worksheet_schedule.update_value('D1', 'start time')
    worksheet_schedule.update_value('E1', 'end time')
    worksheet_schedule.update_value('F1', 'time string')

    return

# Function 3: 將 Gsheet 資料下載至csv
# 主程式沒有呼叫，寫在 def ClickBtnResult(self) 的最前面  
def read_GSheet_to_csv(meeting_name, cred_path, output_csv):
    gc = pygsheets.authorize(service_file = cred_path)  
    sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1h9r2dgUJD5Fze36lkCy0sbpeQAw4V0PJlcc0cVLU4iw/edit#gid=0')

    # 下載該指定分頁名稱的資料，至指定路徑的csv (可新建or覆蓋沿用)
    worksheet = sheet.worksheet('title', meeting_name)
    df = worksheet.get_as_df()
    df.to_csv(output_csv , errors='replace', encoding='utf_8_sig', index=False)  # utf_8_sig  在用import codecs 前，big5 OK!
    
    return


### 變數區域：路徑依據 windows 拼出絕對路徑格式
script_path = os.path.abspath(__file__)  # 取得當前腳本的絕對路徑
current_dir = os.path.dirname(script_path)  # 取得當前腳本所在的資料夾路徑

# Function 2-1 # def clickBtnSubmit(self):
meeting_name = '05/25 Meeting'
cred_path =  './Google Sheet/credentials_ServiceAccount.json'
input_csv = './Google Calendar API/schedule.csv' 
# Function 2-2
sheet_name = 'Overview'  # this is fixed
input_meeting = './meeting.csv'
# Function 3
output_csv = './schedule_output.csv'
#Function 4
return_date = sendback_date(meeting_name)


"""
### 主要程式碼
# Function 2-1: Attendee 填入時間 / Attendee 按下button後的執行動作
# attendee_action = write_csv_to_GSheet(meeting_name, cred_path, input_csv)

# OK # Function 2-2: Host 填入會議資訊 / Host 按下button後的執行動作   
host_action_step1 = write_meeting_csv_to_GSheet(sheet_name, cred_path, input_meeting)

# Function 3: Host 查看最終開會時間
host_action_step2 = read_GSheet_to_csv(meeting_name, cred_path, output_csv)
"""




### 備份 ###
# 下載該指定分頁的資料
# df = worksheet.get_as_df()
# df.to_csv(r'D:\Python_111-2\01_G sheet API\DT_download.csv', errors='replace', encoding='utf-8_sig', index=False)
# worksheet.export(pygsheets.ExportType.CSV)

# 中文亂碼處理
# worksheet.export(file_format=pygsheets.ExportType.CSV, filename='download_test', path='D:/Python_111-2/01_G sheet API/')
