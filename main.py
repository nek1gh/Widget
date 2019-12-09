import httplib2
import apiclient.discovery
from pprint import pprint
from oauth2client.service_account import ServiceAccountCredentials
from urllib.parse import urlparse

CREDENTIALS_FILE = 'credentials.json'
SCOPE = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file',
         'https://www.googleapis.com/auth/drive']
URL = 'https://docs.google.com/spreadsheets/d/1C62PRaDPEwFuyRRZ1XRugPbQIOTZ2pZsMRvBjGO4OM0/edit#gid=0'
arr = URL.split('/')
spreadsheet_id = arr[5]
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPE)
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)
# Google Sheets reading examples
values_read = service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id,
    range='A1:E10',
    majorDimension='ROWS'
).execute()
# Google Sheets writing examples
values_write = service.spreadsheets().values().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body={
        'valueInputOption': 'USER_ENTERED',
        'data': [{
            'range': 'B3:C4',
            'majorDimension': 'ROWS',
            "values": [["This is B2", "This is C2"], ["This is B3", "This is C3"]]},
            {"range": "D5:E6",
             "majorDimension": "COLUMNS",
             "values": [["This is D5", "This is D6"], ["This is E5", "=5+5"]]}
        ]
    }).execute()
