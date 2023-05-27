# 05/27 將 csv 與 GSheet 上傳/下載寫成函數 (測試請自行取消註解)
import google.auth
from google.oauth2.service_account import Credentials
import pandas as pd
import pygsheets
import gspread

# 將 csv 資料延續貼到 Gsheet
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
    
# 將 Gsheet 資料下載至csv
def read_GSheet_to_csv(meeting_name, cred_path, output_csv):
    gc = pygsheets.authorize(service_file = cred_path)  
    sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1h9r2dgUJD5Fze36lkCy0sbpeQAw4V0PJlcc0cVLU4iw/edit#gid=0')

    # 下載該指定分頁名稱的資料，至指定路徑的csv (可新建or覆蓋沿用)
    worksheet = sheet.worksheet('title', meeting_name)
    df = worksheet.get_as_df()
    df.to_csv(output_csv , errors='replace', encoding='utf-8_sig', index=False)
    
    return


"""
# Function 2: attendee 填入時間 / Attendee 按下button後的執行動作
# 須改變以下三列字串變數
meeting_name = '05/25 Meeting'
cred_path = 'D:\\Python_111-2\\01_G sheet API\\credentials.json'
input_csv = "D:\\Python_111-2\\01_G sheet API\\DT_upload.csv" # sample CSV file, use your path

attendee_action = write_csv_to_GSheet(meeting_name, cred_path, csv_path)



# Function 3: 查看最終開會時間
# 須改變以下三列字串變數
meeting_name = '05/25 Meeting'
cred_path = 'D:\\Python_111-2\\01_G sheet API\\credentials.json'
output_csv = 'D:\\Python_111-2\\01_G sheet API\\DT_download.csv'

host_action = read_GSheet_to_csv(meeting_name, cred_path, output_csv)
"""


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