# 05/24 將csv資料寫入G Sheet
import google.auth
from google.oauth2.service_account import Credentials
import pandas as pd
import pygsheets

# 授權金鑰 json 放置的位子
gc = pygsheets.authorize(service_file='D:/Python_111-2/01_G sheet API/credentials.json')  
# 利用 pygsheets 開啟 GoogleSheet
sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1h9r2dgUJD5Fze36lkCy0sbpeQAw4V0PJlcc0cVLU4iw/edit#gid=0')

# 創建指定命名的分頁，並將開會時間寫入該指定分頁
worksheet = sheet.add_worksheet("05/25 new sheet")
worksheet.clear()
df = pd.read_csv('D:/Python_111-2/01_G sheet API/schedule.csv', encoding = 'utf-8')
worksheet.set_dataframe(df,(1,1))

#import os
#os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "D:/Python_111-2/01_G sheet API/credentials.json" 
#creds, _ = google.auth.default()

# scope = ['https://www.googleapis.com/auth/spreadsheets']
# creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
#gs = gspread.authorize(creds)
#sheet = gs.open_by_url('https://docs.google.com/spreadsheets/d/1h9r2dgUJD5Fze36lkCy0sbpeQAw4V0PJlcc0cVLU4iw/edit?usp=sharing')

# worksheet = sheet.worksheet('title', 'create new sheet')


# worksheet.update([df.columns.values.tolist()] + df.values.tolist())




"""
from __future__ import print_function

import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import os
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "D:/Python_111-2/01_G sheet API/credentials.json" 
creds, _ = google.auth.default()

# Create Spreadsheet to the root folder.
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()
body = {}
results = sheet.create(body=body).execute()
print(results)
        
# Move the created Spreadsheet to the specific folder.
drive = build('drive', 'v3', credentials=creds)
folderId = '1s4_-WJhNnkOnzO7dbJ6FUPOKVOyx1MR6' #'### folderId ###'
res = drive.files().update(fileId=results['spreadsheetId'], addParents=folderId, removeParents='root').execute()
print(res)

"""



"""
from pprint import pprint
import google.auth
from googleapiclient import discovery

import os
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "D:/Python_111-2/01_G sheet API/credentials.json" 
# "path_to_your_.json_credential_file"


# TODO: Change placeholder below to generate authentication credentials. See
# https://developers.google.com/sheets/quickstart/python#step_3_set_up_the_sample
#
# Authorize using one of the following scopes:
#     'https://www.googleapis.com/auth/drive'
#     'https://www.googleapis.com/auth/drive.file'
#     'https://www.googleapis.com/auth/spreadsheets'

credentials = None  # 不是字串路徑 'D:/Python_111-2/01_G sheet API/credentials.json'
# creds, _ = google.auth.default()
# credentials = creds
service = discovery.build('sheets', 'v4', credentials=credentials)

spreadsheet_body = {
    # TODO: Add desired entries to the request body.
}

request = service.spreadsheets().create(body=spreadsheet_body)
response = request.execute()

# TODO: Change code below to process the `response` dict:
pprint(response)

"""





"""
from __future__ import print_function

import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import os
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "D:/Python_111-2/01_G sheet API/credentials.json" 

def create(title):
    
    # Creates the Sheet the user has access to.
    # Load pre-authorized user credentials from the environment.
    # TODO(developer) - See https://developers.google.com/identity
    # for guides on implementing OAuth2 for the application.
       
    creds, _ = google.auth.default()
    # pylint: disable=maybe-no-member
    try:
        service = build('sheets', 'v4', credentials=creds)
        spreadsheet = {
            'properties': {
                'title': title
            }
        }
        spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                                    fields='spreadsheetId') \
            .execute()
        print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
        
        return spreadsheet.get('spreadsheetId')
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


if __name__ == '__main__':
    # Pass: title
    create("mysheet1")
"""


"""
# 也有存不到credential.json
from __future__ import print_function

import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


def create(title):
    
    # Creates the Sheet the user has access to.
    # Load pre-authorized user credentials from the environment.
    # TODO(developer) - See https://developers.google.com/identity
    # for guides on implementing OAuth2 for the application.
        # 
    creds, _ = google.auth.default()
    # pylint: disable=maybe-no-member
    try:
        service = build('sheets', 'v4', credentials=creds)
        spreadsheet = {
            'properties': {
                'title': title
            }
        }
        spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                                    fields='spreadsheetId') \
            .execute()
        print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
        return spreadsheet.get('spreadsheetId')
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


if __name__ == '__main__':
    # Pass: title
    create("mysheet1")
"""



"""
# 可能要在quickstart 取得新的token後才能順利跑，但目前quickstart 遇到credentials.json 抓取問題

# BEFORE RUNNING:
# ---------------
# 1. If not already done, enable the Google Sheets API
   # and check the quota for your project at
   # https://console.developers.google.com/apis/api/sheets
# 2. Install the Python client library for Google APIs by running
   # `pip install --upgrade google-api-python-client`


from pprint import pprint
from googleapiclient import discovery

# TODO: Change placeholder below to generate authentication credentials. See
# https://developers.google.com/sheets/quickstart/python#step_3_set_up_the_sample
#
# Authorize using one of the following scopes:
#     'https://www.googleapis.com/auth/drive'
#     'https://www.googleapis.com/auth/drive.file'
#     'https://www.googleapis.com/auth/spreadsheets'
credentials = None

service = discovery.build('sheets', 'v4', credentials=credentials)

spreadsheet_body = {
    # TODO: Add desired entries to the request body.
}

request = service.spreadsheets().create(body=spreadsheet_body)
response = request.execute()

# TODO: Change code below to process the `response` dict:
pprint(response)
"""




"""
# 根據既有工作表編輯內容
import pygsheets
import pandas as pd

gc = pygsheets.authorize(service_file='D:/Python_111-2/01_G sheet API/credentials.json')  # 授權金鑰 json 放置的位子

# 利用 Python 開啟 GoogleSheet
sht = gc.open_by_url(
'https://docs.google.com/spreadsheets/d/1tlD0m4tZSL10vs30ds5Rjia5_Oa-JdTHAeKEhBtTKRk/edit#gid=0'
)

# 查看此 GoogleSheet 內 Sheet 清單
wks_list = sht.worksheets()
# print(wks_list)

#選取 Sheet by順序
wks = sht[0]
#選取Sheet by名稱
# wks2 = sht.worksheet_by_title("Sheet2")
 
# 更新名稱  
# 依據會議名稱更新 tab/ sheet name

wks.title = "0523 demo Title"
#隱藏清單
wks.hidden = False

# Create empty dataframe
df = pd.DataFrame()
# Create a column
df['name'] = ['John', 'Steve', 'Sarah']

#update the first sheet with df, starting at cell B2. 
wks.set_dataframe(df,(1,1))
"""



"""
# 參考網友解方，已安裝額外套件 npm install 
# https://ithelp.ithome.com.tw/articles/10234325

// googleSheet.js

const { GoogleSpreadsheet } = require('google-spreadsheet');

/**
 * @param  {String} docID the document ID
 * @param  {String} sheetID the google sheet table ID
 * @param  {String} credentialsPath the credentials path defalt is './credentials.json'
 */
async function getData(docID, sheetID, credentialsPath = './credentials.json') {
  const result = [];
  const doc = new GoogleSpreadsheet(docID);
  const creds = require(credentialsPath);
  await doc.useServiceAccountAuth(creds);
  await doc.loadInfo();
  const sheet = doc.sheetsById[sheetID];
  const rows = await sheet.getRows();
  for (row of rows) {
    result.push(row._rawData);
  }
  return result;
};

module.exports = {
  getData,
};





// test.js
const { getData } = require('./googleSheet.js');

(async () => {
  const resp = await getData('<docID>', '<sheetID>');
  console.log(resp);
})();
"""