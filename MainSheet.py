from typing import Union

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet

from BorderCoord import BorderCoord as bc
from Coord import Coord


class MainSheet:

    def __init__(self, wb: Workbook, subjects: dict):
        self.subjects = subjects
        self.wb = wb
        self.create_sheet()

    def create_sheet(self):
        ws = self.wb.create_sheet("Ведомость усп.и посещ.")

        self.create_title(ws)

    def create_title(self, ws: Union[WriteOnlyWorksheet, Worksheet]):
        nums_row_vertical = [3, 4]

        title_left = {
            "№ п/п": Coord([nums_row_vertical[0], 1], [nums_row_vertical[1], 1],
                           border=bc.get_border_thin(top=bc.MEDIUM, left=bc.MEDIUM)),
            "ФИО": Coord([nums_row_vertical[0], 2], [nums_row_vertical[1], 2],
                         border=bc.get_border_thin(top=bc.MEDIUM))
        }
        self.draw_all(ws, title_left, font_style_bold=True)

        title_bottom_middle = dict()
        last_col = 2
        for title_subject in self.subjects:
            last_col += 1
            title_bottom_middle[title_subject] = Coord([nums_row_vertical[1], last_col], border=bc.get_border_thin())
        self.draw_all(ws, title_bottom_middle, font_size=Coord.FONT_SIZE_MINI, font_style_bold=True)

        title_top_middle = {"Дисциплина": Coord([nums_row_vertical[0], 3], [nums_row_vertical[0], last_col],
                                                border=bc.get_border_thin(top=bc.MEDIUM))}
        self.draw_all(ws, title_top_middle, font_style_bold=True)

        last_col += 1
        title_right = {"Сред. балл": Coord([nums_row_vertical[0], last_col], [nums_row_vertical[1], last_col],
                                           border=bc.get_border_medium(bottom=bc.THIN))}
        self.draw_all(ws, title_right, font_style_bold=True)

        last_col += 1
        right_title_top = {
            "Пропущено часов": Coord([nums_row_vertical[0], last_col], [nums_row_vertical[0], last_col + 2],
                                     border=bc.get_border_medium(bottom=bc.THIN))}
        self.draw_all(ws, right_title_top, font_style_bold=True)

        right_title_bottom = {
            "Всего": Coord([nums_row_vertical[1], last_col], border=bc.get_border_thin(left=bc.MEDIUM)),
            "По уваж.прич.": Coord([nums_row_vertical[1], last_col + 1], border=bc.get_border_thin()),
            "По не уваж.прич.": Coord([nums_row_vertical[1], last_col + 2], border=bc.get_border_thin(right=bc.MEDIUM)),
        }
        self.draw_all(ws, right_title_bottom)
        last_col += 2

        self.wb.close()

    def draw_all(self, ws, d: dict, font=None, font_size=None, font_style_bold: bool = None):
        for title, coord in d.items():
            coord.draw(ws=ws, title=title, font=font, font_size=font_size, font_style_bold=font_style_bold)
