# 若下載後於個人電腦，請自行修改路徑如： credentials.json、meeting.csv、schedule.csv (ex. C:\\your\\own\\path\\file.csv)
import google.auth
from google.oauth2.service_account import Credentials
import pandas as pd
import pygsheets
import gspread

# Function 2-1: 將 schedule.csv 資料延續貼到 Gsheet
def write_csv_to_GSheet(meeting_name, cred_path, input_csv):
    # 利用 pygsheets 開啟 GoogleSheet
    gc = pygsheets.authorize(service_account_file = cred_path)
    sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1h9r2dgUJD5Fze36lkCy0sbpeQAw4V0PJlcc0cVLU4iw/edit#gid=0')
    worksheet = sheet.worksheet('title', meeting_name)

    # sample CSV file, use your path
    df = pd.read_csv(input_csv, encoding = 'utf-8')
    cells = worksheet.get_all_values(include_tailing_empty_rows=False, include_tailing_empty=False, returnas='matrix')
    last_row = str(len(cells)+1)
    worksheet.set_dataframe(df, start = "A" + last_row ,copy_head=False)

    return
    

# Function 2-2: 將 meeting.csv 資料延續貼到 Gsheet
# 並在 schedule 的 spreadsheet 根據會議名稱建立新的分頁
def write_meeting_csv_to_GSheet(sheet_name, cred_path, input_meeting):
    # 利用 pygsheets 開啟 GoogleSheet
    gc = pygsheets.authorize(service_account_file = cred_path)
    sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/18THYdAFqANCRORZkBp7tDgGgjXGeIWdLFsJK_0ez0qM/edit#gid=0')
    worksheet = sheet.worksheet('title', sheet_name)

    df = pd.read_csv(input_meeting, encoding = 'utf-8')
    cells = worksheet.get_all_values(include_tailing_empty_rows=False, include_tailing_empty=False, returnas='matrix')
    last_row = str(len(cells)+1)
    worksheet.set_dataframe(df, start = "A" + last_row ,copy_head=False)

    ### 填入title, 抓會議名稱
    sheet_schedule = gc.open_by_url('https://docs.google.com/spreadsheets/d/1h9r2dgUJD5Fze36lkCy0sbpeQAw4V0PJlcc0cVLU4iw/edit#gid=0')
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
def read_GSheet_to_csv(meeting_name, cred_path, output_csv):
    gc = pygsheets.authorize(service_file = cred_path)  
    sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1h9r2dgUJD5Fze36lkCy0sbpeQAw4V0PJlcc0cVLU4iw/edit#gid=0')

    # 下載該指定分頁名稱的資料，至指定路徑的csv (可新建or覆蓋沿用)
    worksheet = sheet.worksheet('title', meeting_name)
    df = worksheet.get_as_df()
    df.to_csv(output_csv , errors='replace', encoding='utf-8_sig', index=False)
    
    return



# Function 2-1: Attendee 填入時間 / Attendee 按下button後的執行動作
# 須改變以下三列字串變數
meeting_name = '05/25 Meeting'
cred_path = './Google Sheet/credentials_ServiceAccount.json'
input_csv = "./schedule.csv" # sample CSV file, use your path

attendee_action = write_csv_to_GSheet(meeting_name, cred_path, input_csv)




# Function 2-2: Host 填入會議資訊 / Host 按下button後的執行動作
# 須改變以下三列字串變數
sheet_name = 'Overview'  # this is fixed
cred_path = './Google Sheet/credentials_ServiceAccount.json'  # 請使用個人路徑
input_meeting ="./meeting.csv"  # 請使用個人路徑

host_action_step1 = write_meeting_csv_to_GSheet(sheet_name, cred_path, input_meeting)



# Function 3: Host 查看最終開會時間
# 須改變以下三列字串變數
meeting_name = '05/25 Meeting'
cred_path = './Google Sheet/credentials_ServiceAccount.json'
output_csv = './schedule_output.csv'

host_action_step2 = read_GSheet_to_csv(meeting_name, cred_path, output_csv)



### 待辦函數 ###
# 創建指定命名的分頁，並將開會時間寫入該指定分頁
# meeting_name = '05/25 Meeting'
# worksheet = sheet.add_worksheet(meeting_name)


### 備份 ###
# 下載該指定分頁的資料
# df = worksheet.get_as_df()
# df.to_csv(r'D:\Python_111-2\01_G sheet API\DT_download.csv', errors='replace', encoding='utf-8_sig', index=False)
# worksheet.export(pygsheets.ExportType.CSV)

# 中文亂碼處理
# worksheet.export(file_format=pygsheets.ExportType.CSV, filename='download_test', path='D:/Python_111-2/01_G sheet API/')
