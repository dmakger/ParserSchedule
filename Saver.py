import datetime
import time

from openpyxl.styles.borders import Border, Side
from openpyxl import Workbook
import os

from MainSheet import MainSheet
from SkipSheet import SkipSheet
from SubjectsSheet import SubjectsSheet


class Saver:
    # FILE_NAME = "test.xlsx"

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

        self.file_name = self.get_file_name()

    def get_file_name(self):
        month = str(self.month)
        if len(month) == 1:
            month = f'0{month}'
        return f'Отчёт {self.group} {month}.{self.year}.xlsx'

    def save(self):
        wb = Workbook()
        self.delete_sheets(wb)

        name_skip_sheet = f"{self.group} (ЭН)"
        print(f'Создание листа: "{name_skip_sheet}"')
        skip_sheet = SkipSheet(wb=wb, subjects=self.subjects, schedule=self.schedule, name_sheet=name_skip_sheet,
                               month=self.month, year=self.year, all_days=self.all_days)
        print()
        name_subjects_sheet = f"{self.group} (БН)"
        print(f'Создание листа: "{name_subjects_sheet}"')
        subjects_sheet = SubjectsSheet(wb=wb, subjects=self.subjects, schedule=self.schedule,
                                       name_sheet=name_subjects_sheet,
                                       month=self.month, year=self.year, all_days=self.all_days)
        print()
        print(f'Создание листа: "{MainSheet.NAME_SHEET}"')
        MainSheet(wb=wb, subjects=self.subjects, speciality=self.speciality, group=self.group, month=self.month, data={
            'skip': {
                'name_sheet': name_skip_sheet,
                'last_row': skip_sheet.last_row,
                'col_end': skip_sheet.col_end,
                'result_start_row': skip_sheet.result_start_row
            },
            'subjects': {
                'name_sheet': name_subjects_sheet,
                'last_row': subjects_sheet.last_row,
                'col_end': subjects_sheet.col_end,
                'result_start_row': subjects_sheet.result_start_row,
                'result_start_col': subjects_sheet.result_start_col
            }
        })
        print()

        wb.active = 2
        print("Сохранение файла...")
        wb.save(self.file_name)
        print("Открываем файл")
        os.startfile(self.file_name)

    def delete_sheets(self, wb):
        for sheet_name in wb.sheetnames:
            sheet = wb.get_sheet_by_name(sheet_name)
            wb.remove_sheet(sheet)
