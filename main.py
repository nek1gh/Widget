import json
import datetime
import httplib2
import apiclient.discovery
import googleapiclient.errors
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
d = 0
for i in leads:
    row_data = dict()
    row_data['id'] = i['id']
    row_data['status_id'] = status_all[i['status_id']]
    row_data['company_id'] = i['company']['id'] if i['company'] != {} else ""
    row_data['company_name'] = i['company']['name'] if i['company'] != {} else ""
    row_data['created_at'] = str(change_date(i['created_at']))
    row_data['updated_at'] = str(change_date(i['updated_at']))
    # pprint(row_data)
    d += len(row_data)

CREDENTIALS_FILE = 'credentials.json'
SCOPE = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file',
         'https://www.googleapis.com/auth/drive']
URL = 'https://docs.google.com/spreadsheets/d/1C62PRaDPEwFuyRRZ1XRugPbQIOTZ2pZsMRvBjGO4OM0/edit#gid=0'
arr = URL.split('/')
spreadsheet_id = arr[5]
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPE)
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)
header = ['ID', 'Имя статуса', 'ID компании', 'Имя компании', 'Дата создания', 'Дата завершения', 'Дата Отгрузки']
# Google Sheets writing examples
values_write = service.spreadsheets().values().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body={
        'valueInputOption': 'USER_ENTERED',
        'data': [{
            'range': 'A1:H1',
            'majorDimension': 'ROWS',
            "values": [[header[0], header[1], header[2], header[3], header[4], header[5], header[6]]]}
        ]
    }).execute()

# Google Sheets reading examples
'''values_read = service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id,
    range='A1:E10',
    majorDimension='ROWS'
).execute()'''
