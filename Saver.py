import datetime
import time

from openpyxl.styles.borders import Border, Side
from openpyxl import Workbook
import os

from MainSheet import MainSheet
from SkipSheet import SkipSheet


class Saver:
    FILE_NAME = "test.xlsx"
    
    def __init__(self, schedule, subjects, group, speciality, all_days, month=None, year=None):
        self.schedule = schedule
        self.subjects = subjects
        self.group = group
        self.speciality = speciality
        self.all_days = all_days

        if month is None:
            month = datetime.date.today().month
        self.month = month

        if year is None:
            year = datetime.date.today().year
        self.year = year

        self.save()

    def save(self):
        wb = Workbook()
        self.delete_sheets(wb)

        print(f"Создание листа: {MainSheet.NAME_SHEET}...")
        MainSheet(wb=wb, subjects=self.subjects, speciality=self.speciality, group=self.group, month=self.month)
        name_two_sheet = f"{self.group} (ЭН)"
        print(f"Создание листа: {name_two_sheet}...")
        SkipSheet(wb=wb, subjects=self.subjects, schedule=self.schedule, name_sheet=name_two_sheet,
                  month=self.month, year=self.year, all_days=self.all_days)

        wb.active = 1
        wb.save(self.FILE_NAME)
        os.startfile(self.FILE_NAME)

    def delete_sheets(self, wb):
        for sheet_name in wb.sheetnames:
            sheet = wb.get_sheet_by_name(sheet_name)
            wb.remove_sheet(sheet)




