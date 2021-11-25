import time

from openpyxl.styles.borders import Border, Side
from openpyxl import Workbook
import os

from MainSheet import MainSheet


class Saver:
    FILE_NAME = "test.xlsx"
    
    def __init__(self, schedule, subjects, group, month=None):
        self.schedule = schedule
        self.subjects = subjects
        self.group = group
        self.month = month
        self.save()

    def save(self):
        wb = Workbook()
        self.delete_sheets(wb)

        main_sheet = MainSheet(wb, self.subjects)

        # self.create_main(wb)
        # ws = wb.active
        # ws.title = "Ведомость усп.и посещ."
        # thin_border = Border(left=Side(style='thin'),
        #                      right=Side(style='thin'),
        #                      top=Side(style='thin'),
        #                      bottom=Side(style='thin'))
        #
        # # property cell.border should be used instead of cell.style.border
        # ws.cell(row=3, column=2).border = thin_border
        wb.save(self.FILE_NAME)
        os.startfile(self.FILE_NAME)

    def delete_sheets(self, wb):
        for sheet_name in wb.sheetnames:
            sheet = wb.get_sheet_by_name(sheet_name)
            wb.remove_sheet(sheet)

    # def get_title_month(self):
    def create_main(self, wb):
        ws = wb.create_sheet("Ведомость усп.и посещ.")
        # ws.merge_cells()

        title_left = ["№ п/п", "ФИО"]
        title_top_middle = ["Дисциплина"]
        title_bottom_middle = []
        title_right = ["Сред. балл"]
        right_title_top = ["Пропущено часов"]
        right_title_bottom = ["Всего", "По уваж.прич.", "По не уваж.прич."]

        for title_subject in self.subjects:
            title_bottom_middle.append(title_subject)




