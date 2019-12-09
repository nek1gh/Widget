import json
import datetime
import httplib2
import apiclient.discovery
from pprint import pprint
from oauth2client.service_account import ServiceAccountCredentials


# From number to date
def change_date(timestamp):
    return datetime.datetime.utcfromtimestamp(timestamp // 24 // 60 // 60 * 60 * 60 * 24)


# Parse JSON status
sJSON = 'status.json'
dataStatus = json.load(open(sJSON))
status = dataStatus['_embedded']['items']
status_all = {status[funnel]['statuses'][i]['id']: status[funnel]['statuses'][i]['name'] for funnel in status for i
              in status[funnel]['statuses']}
# Parse JSON leads
JSON = 'amocrm.json'
data = json.load(open(JSON))
leads = data['_embedded']['items']
for i in leads[:1]:
    row_data = dict()
    row_data['id'] = i['id']
    row_data['status_id'] = status_all[i['status_id']]
    row_data['company_id'] = i['company']['id'] if i['company'] != {} else ""
    row_data['company_name'] = i['company']['name'] if i['company'] != {} else ""
    row_data['created_at'] = str(change_date(i['created_at']))
    row_data['updated_at'] = str(change_date(i['updated_at']))
    pprint(row_data)

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
