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


