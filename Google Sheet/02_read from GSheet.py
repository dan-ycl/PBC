# 從 G Sheet下載完整 cdv
import google.auth
from google.oauth2.service_account import Credentials
import pandas as pd
import pygsheets

# 授權金鑰 json 放置的位子
gc = pygsheets.authorize(service_file='D:/Python_111-2/01_G sheet API/credentials.json')  
# 利用 pygsheets 開啟 GoogleSheet
sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1h9r2dgUJD5Fze36lkCy0sbpeQAw4V0PJlcc0cVLU4iw/edit#gid=0')


# 下載該指定分頁名稱的資料，至指定路徑的csv (可新建or覆蓋沿用)
worksheet = sheet.worksheet('title', '05/25 new sheet')
df = worksheet.get_as_df()
df.to_csv(r'D:\Python_111-2\01_G sheet API\AAA.csv', errors='replace', encoding='utf-8_sig', index=False)

