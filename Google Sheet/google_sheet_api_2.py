import os
from Google import Create_Service

FOLDER_PATH = r'<Folder Path>'
CLIENT_SECRET_FILE = os.path.join(r'C:\Users\maggy\OneDrive\文件\GitHub\PBC\Google Sheet', 'Client_Secret.json')
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)