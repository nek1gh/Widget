import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


class Spreadsheet:
    def __init__(self, json_key):
        SCOPE = ['https://www.googleapis.com/auth/spreadsheets',
                 'https://www.googleapis.com/auth/drive']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(json_key, SCOPE)
        self.httpAuth = self.credentials.authorize(httplib2.Http())
        self.service = apiclient.discovery.build('sheets', 'v4', http=self.httpAuth)
        self.spreadsheetId = None  # ID таблицы
        self.sheetTitle = None  # Название листа
        self.sheetId = None  # ID листа
        self.data = []  # Данные передаваемые по API
        self.requests = []

    def set_spreadsheet_by_id(self, spreadsheetId):
        spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
        self.spreadsheetId = spreadsheet['spreadsheetId']
        self.sheetTitle = spreadsheet['sheets'][0]['properties']['title']
        self.sheetId = spreadsheet['sheets'][0]['properties']['sheetId']

    def gridRange(self, cellsRange):
        if isinstance(cellsRange, str):  # Проверяет что переменная является строкой
            startCell, endCell = cellsRange.split(":")[0:2]  # Выясняем начальную и конечную ячейку
            cellsRange = {}  # Пустой словарь для диапазона ячеек
            # Возвращает диапазон символа по таблице Unicode.
            rangeAZ = range(ord('A'), ord('Z') + 1)  # A = 65, Z = 90 + 1. Т.к. последнее значение не учитывается.
            #  В ходил ли эта буква в диапазон от 65 до 91
            if ord(startCell[0]) in rangeAZ:
                # Записывает в словарь ключ со значением начала столбца
                cellsRange["startColumnIndex"] = ord(startCell[0]) - ord('A')  # startCell[0] - 65
                startCell = startCell[1:]  # Записывает в переменную значение после буквы
            #  В ходил ли эта буква в диапазон от 65 до 91
            if ord(endCell[0]) in rangeAZ:
                # Записывает в словарь ключ со значением окончания столбца
                # Увеличено на единицу, т.к. последнее значение не учитывается.
                cellsRange["endColumnIndex"] = ord(endCell[0]) - ord('A') + 1
                endCell = endCell[1:]  # Записывает в переменную значение после буквы
            # Указано ли второе значение после буквы в переменной cellsRange
            if len(startCell) > 0:
                cellsRange["startRowIndex"] = int(startCell) - 1  # Вычитаем единицу т.к нумерация начинается с нуля
            # Указано ли второе значение после буквы в переменной cellsRange
            if len(endCell) > 0:
                cellsRange["endRowIndex"] = int(endCell)  # Оставляем т.к. есть
            cellsRange["sheetId"] = self.sheetId  # Записываем в словарь ID листа
            return cellsRange
            # cell_range = {'endColumnIndex': ...,
            #               'endRowIndex': ...,
            #               'sheetId': ...,
            #               'startColumnIndex': ...,
            #               'startRowIndex': ...}

    def runBatch(self, valueInputOption="USER_ENTERED"):
        replies = {'replies': []}
        responses = {'responses': []}
        try:
            if len(self.requests) > 0:
                replies = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId,
                                                                  body={"requests": self.requests}).execute()
            if len(self.data) > 0:
                responses = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId,
                                                                             body={'valueInputOption': valueInputOption,
                                                                                   'data': self.data}).execute()
        finally:
            self.data = []
            self.requests = []
        return replies['replies'], responses['responses']

    def repeat_cell(self, cellsRange, userEF, fields="userEnteredFormat"):
        self.requests.append(
            {"repeatCell": {'range': self.gridRange(cellsRange), 'cell': {'userEnteredFormat': userEF},
                            'fields': fields}})

    def set_values(self, cellsRange, values, majorDimension="ROWS"):
        self.data.append(
            {"range": self.sheetTitle + "!" + cellsRange, "majorDimension": majorDimension, "values": values})
