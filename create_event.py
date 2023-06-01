from datetime import datetime, timedelta
from cal_setup import get_calendar_service
from uuid import uuid4

def create_event(meeting_name, start, end, CREDENTIALS_FILE, attend_list):
   
   service = get_calendar_service(CREDENTIALS_FILE)

   event_result = service.events().insert(calendarId='primary',
       body={
           "summary": meeting_name,
           # "description": 'This is a tutorial example of automating google calendar with python',
           "start": {"dateTime": start, "timeZone": 'Asia/Taipei'},
           "end": {"dateTime": end, "timeZone": 'Asia/Taipei'},
           "attendees":attend_list,
           "conferenceData": {"createRequest": {"requestId": f"{uuid4().hex}",
                                                  "conferenceSolutionKey": {"type": "hangoutsMeet"}}},    
           },
       sendUpdates = 'all', conferenceDataVersion=1
   ).execute()

   # print 部分可以省略，單純在 terminal上顯示訊息
   print("created event")
   print("id: ", event_result['id'])
   print("summary: ", event_result['summary'])
   print("starts at: ", event_result['start']['dateTime'])
   print("ends at: ", event_result['end']['dateTime'])

# 必須是 oauth 的jason，和google sheet 的金鑰(屬於service account)不一樣
# https://console.cloud.google.com/apis/credentials?project=festive-catwalk-387114 或於此網址下載(Google_OAuth_Calendar.json)



# CREDENTIALS_FILE = './Google Calendar API/credentials_oauth.json' 
# meeting_name = '05/25 Meeting'
# # 測試版的會議時間：creates one hour event tomorrow 10 AM IST
# d = datetime.now().date()
# tomorrow = datetime(d.year, d.month, d.day, 10)+timedelta(days=1)
# start = tomorrow.isoformat()
# end = (tomorrow + timedelta(hours=2)).isoformat()


# print(start)
# print(end)

# # email 參與者 (小奇測試個人email帳號 都ok)
# email_list = ['hjuo37689@gmail.com']
# attend_list = [0] * len(email_list)

# for i in range(len(email_list)):
#     attend_list[i] = {'email': email_list[i]}

# # Function 1：host 發起會議
# # 需要下載 cal_setup.py 、 oauth權限的 crendential.json 和 pip google oauth
# # cal_setup.py 請和create.py放在同一資料夾
# # 共須傳入五個參數
# host_setup = create_event(meeting_name, start, end, CREDENTIALS_FILE, attend_list)
