import calendar
import locale
import os
from datetime import datetime, date

import gspread
from gspread.exceptions import WorksheetNotFound
from gspread_formatting import *


locale.setlocale(locale.LC_ALL, 'ru_RU.utf-8')  # Переводим все на русский

SERVICE_ACC = os.path.join(os.getcwd(), 'static', 'service_acc.json')
SPREADSHEET_ID_ORDERS = os.getenv('SPREADSHEET_ID_ORDERS')
SPREADSHEET_ID_TIMETABLE = os.getenv('SPREADSHEET_ID_TIMETABLE')


class Dates:

    def __init__(self, month: int, year: int = 2021, day: int = 1):
        self.work_date = datetime(year=year, month=month, day=day)
        self.first_and_count_days = calendar.monthrange(self.work_date.year, self.work_date.month)
        self.year = year
        self.month = month

    def count_day_in_month(self):
        return self.first_and_count_days[1]

    def calendar(self):
        return calendar.monthcalendar(self.work_date.year, self.work_date.month)

    def date_str(self):
        return self.work_date.strftime('%B.%Y')

    def name_days(self):
        first_day_month, count_day = self.first_and_count_days
        weekdays = list(calendar.day_name)
        func_iterator = 0
        while True:
            yield weekdays[first_day_month]
            first_day_month += 1
            func_iterator += 1
            if first_day_month > 6:
                first_day_month = 0
            if func_iterator >= count_day:
                break

    def number_day_is_weekend(self):
        for day in range(1, self.count_day_in_month()):
            if date.weekday(date(self.year, self.month, day)) in [5, 6]:
                yield day

    def __str__(self):
        return f'{self.month}.{self.year}'


FORMAT_HEADER = CellFormat(
    backgroundColor=Color(1, 0.67, 0.5),
    textFormat=TextFormat(italic=True, fontSize=16, foregroundColor=Color(0, 0, 0))
)

FORMAT_BODY = CellFormat(
    backgroundColor=Color(.9, .9, .9),
    textFormat=TextFormat(bold=True, fontSize=12, foregroundColor=Color(0, 0, 0)),
    horizontalAlignment='CENTER',
    verticalAlignment='MIDDLE',
    borders=Borders(Border(style='DOUBLE'), Border(style='DOUBLE'),
                    Border(style='SOLID'), Border(style='SOLID'))
)

FORMAT_WEEKEND = CellFormat(
    backgroundColor=Color(.8, .1, .1),
)


class NewSpreadsheets:

    def __init__(self, spreadsheet_id: str):
        self.spreadsheets = self.auth(key=spreadsheet_id)
        self.title = None

    def auth(self, key: str):
        return gspread.service_account(SERVICE_ACC).open_by_key(key=key)

    def get_value(self, sheet_title: str):
        try:
            response = self.spreadsheets.worksheet(title=sheet_title).get_all_values()
        except WorksheetNotFound as exc:
            response = [[f'NotFoundWorksheetWithTitle - {exc}', ], ]
        return response

    def __str__(self):
        return f'{self.spreadsheets.url}\nspreadsheet_title - {self.spreadsheets.title}\nid - {self.spreadsheets.id}'


if __name__ == '__main__':
     test = NewSpreadsheets(SPREADSHEET_ID_TIMETABLE)
