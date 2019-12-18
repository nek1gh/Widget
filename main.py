import json
import copy
import datetime
from pprint import pprint
from googlesheets import Spreadsheet

CREDENTIALS_FILE = 'credentials.json'


def htmlColorToJSON(htmlColor):
    if htmlColor.startswith("#"):
        htmlColor = htmlColor[1:]
    return {"red": int(htmlColor[0:2], 16) / 255.0, "green": int(htmlColor[2:4], 16) / 255.0,
            "blue": int(htmlColor[4:6], 16) / 255.0}


# From number to date
def change_date(timestamp):
    date = str(datetime.datetime.utcfromtimestamp(timestamp // 24 // 60 // 60 * 60 * 60 * 24)).split()
    return date[0]


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
row_dataS = []
a = []
for i in leads:
    row_data = []
    row_data.append(i['id'])
    row_data.append(status_all[i['status_id']])
    row_data.append(i['company']['id'] if i['company'] != {} else "")
    row_data.append(i['company']['name'] if i['company'] != {} else "")
    row_data.append(str(change_date(i['created_at'])))
    row_data.append(str(change_date(i['closed_at']) if i['closed_at'] != 0 else ""))
    a += [i['custom_fields']]
    count = 0
    tovars = {}
    n_rejsa = ""
    for j in range(len(i['custom_fields'])):
        name = i['custom_fields'][j]["name"]
        if "Товар" in name:
            nazv = i['custom_fields'][j]["values"][0]["value"] if "value" in i['custom_fields'][j]["values"][0] else ""
            tovars[name] = [nazv,
                            i['custom_fields'][j + 1]["values"][0]["value"],
                            i['custom_fields'][j + 2]["values"][0]["value"]]
        if "№ Рейса" in name:
            n_rejsa = i['custom_fields'][j]["values"][0]["value"]
    row_data.append(n_rejsa)
    if len(tovars) == 0:
        continue
    for tovar in tovars:
        row_data_many = copy.deepcopy(row_data)
        for val in tovars[tovar]:
            row_data_many.append(val)
        row_dataS += [row_data_many]
row_count = len(row_dataS)

URL = 'https://docs.google.com/spreadsheets/d/1YrK6ZXvOBhwP2JqiaE3NzZVv_OpdaKzbyIYcma2j9rk/edit?usp=sharing'
arr = URL.split('/')
spreadsheet_id = arr[5]

header = ['ID', 'Имя статуса', 'ID компании', 'Имя компании', 'Дата создания', 'Дата завершения', 'Дата Отгрузки',
          'Бюджет', 'Вес Заказа, кг', '№ Рейса', 'Населённый пункт', 'Адрес Доставки', 'Товар', 'Цена за упаковку',
          'Количесво']
range_name = "A1:O1"
ss = Spreadsheet(CREDENTIALS_FILE)
ss.set_spreadsheet_by_id(spreadsheet_id)
ss.repeat_cell(range_name,
               {"textFormat": {"bold": True, "foregroundColor": htmlColorToJSON("#FFFFFF"), "fontSize": 11},
                "horizontalAlignment": "CENTER",
                "wrapStrategy": "LEGACY_WRAP",
                "backgroundColor": htmlColorToJSON("#4A8CC5")},
               "userEnteredFormat.textFormat, userEnteredFormat.horizontalAlignment, userEnteredFormat.backgroundColor,"
               "userEnteredFormat.wrapStrategy")
ss.set_values(range_name, [header])

ss.runBatch()
i = 0

ss.set_values("A2:J%d" % (row_count + 2), row_dataS)
ss.repeat_cell("E2:F%d" % (row_count + 2), {"numberFormat": {'pattern': 'dd.mm.yyyy', 'type': 'DATE'}},
               "userEnteredFormat.numberFormat")
ss.runBatch()
