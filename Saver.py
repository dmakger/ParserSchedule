import time

from openpyxl.styles.borders import Border, Side
from openpyxl import Workbook
import os

from MainSheet import MainSheet


class Saver:
    FILE_NAME = "test.xlsx"
    
    def __init__(self, schedule, subjects, group, speciality, month=None):
        self.schedule = schedule
        self.subjects = subjects
        self.group = group
        self.speciality = speciality
        self.month = month
        self.save()

    def save(self):
        wb = Workbook()
        self.delete_sheets(wb)

        MainSheet(wb=wb, subjects=self.subjects, speciality=self.speciality, group=self.group, month=self.month)

        wb.save(self.FILE_NAME)
        os.startfile(self.FILE_NAME)

    def delete_sheets(self, wb):
        for sheet_name in wb.sheetnames:
            sheet = wb.get_sheet_by_name(sheet_name)
            wb.remove_sheet(sheet)




