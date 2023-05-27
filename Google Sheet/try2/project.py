
from googleapiclient.discovery import build
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'keys.json'

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID spreadsheet.
SAMPLE_SPREADSHEET_ID = '1tlD0m4tZSL10vs30ds5Rjia5_Oa-JdTHAeKEhBtTKRk'

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range='0523 demo Title!A1:E6').execute()  # 讀資料
# values = result.get('values', [])


# 想要輸入的資料
put_in = [["05/23/2012", 2000], ["05/24/2012", 5000], ["05/24/2012", 5000]]

# 把要輸入的資料傳到google sheet
request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='Sheet2!B2', valueInputOption="USER_ENTERED", body={"values":put_in}).execute()



