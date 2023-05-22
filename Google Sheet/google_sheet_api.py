from __future__ import print_function

import os.path
import google.auth

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# import os
# os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "$HOME/.config/gcloud/application_default_credentials.json"

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
SAMPLE_RANGE_NAME = 'Class Data!A2:E'


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return

        print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            print('%s, %s' % (row[0], row[4]))
    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()




# def create(title):
#     """
#     Creates the Sheet the user has access to.
#     Load pre-authorized user credentials from the environment.
#     TODO(developer) - See https://developers.google.com/identity
#     for guides on implementing OAuth2 for the application.
#         """
#     creds, _ = google.auth.default()
#     # pylint: disable=maybe-no-member
#     try:
#         service = build('sheets', 'v4', credentials=creds)
#         spreadsheet = {
#             'properties': {
#                 'title': title
#             }
#         }
#         spreadsheet = service.spreadsheets().create(body=spreadsheet,
#                                                     fields='spreadsheetId') \
#             .execute()
#         print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
#         return spreadsheet.get('spreadsheetId')
#     except HttpError as error:
#         print(f"An error occurred: {error}")
#         return error


# if __name__ == '__main__':
#     # Pass: title
#     create("mysheet1")