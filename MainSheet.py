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
        self.wb = wb
        self.subjects = subjects
        self.create_sheet()
        self.last_col = 1

    def create_sheet(self):
        ws = self.wb.create_sheet("Ведомость усп.и посещ.")

        self.create_students(ws)
        self.create_subjects(ws)
        # self.create_statistic(ws)

    def create_students(self, ws: Union[WriteOnlyWorksheet, Worksheet]):
        # Title
        title_left = {
            "№ п/п": Coord([3, 1], [4, 1], border=bc.get_border_thin(top=bc.MEDIUM, left=bc.MEDIUM)),
            "ФИО": Coord([3, 2], [4, 2], border=bc.get_border_thin(top=bc.MEDIUM))
        }
        self.draw_all(ws, title_left, font_style_bold=True)
        self.last_col = 2

        # Main
        students = list(list(self.subjects.values())[0].keys())
        current_row = 5
        for i in range(len(students) - 1):
            Coord([current_row, 1], border=bc.get_border_thin(left=bc.MEDIUM)).draw(ws=ws, title=str(i + 1))
            Coord([current_row, 2], border=bc.get_border_thin()).draw(ws=ws, title=students[i],
                                                                      style_horizontal=Coord.LEFT,
                                                                      )
            current_row += 1
        Coord([current_row, 1], border=bc.get_border_thin(left=bc.MEDIUM, bottom=bc.MEDIUM)).draw(ws=ws,
                                                                                                  title=str(i + 1))
        Coord([current_row, 2], border=bc.get_border_thin(bottom=bc.MEDIUM)).draw(ws=ws, title=students[i],
                                                                                  style_horizontal=Coord.LEFT)

    def create_subjects(self, ws: Union[WriteOnlyWorksheet, Worksheet]):
        # title_bottom_middle = dict()

        print(self.last_col)
        title_bottom_middle = dict()
        # last_col = 2
        # for title_subject in self.subjects:
        #     last_col += 1
        #     title_bottom_middle[title_subject] = Coord([4, last_col], border=bc.get_border_thin())
        # self.draw_all(ws=ws, d=title_bottom_middle, font_size=Coord.FONT_SIZE_MINI, font_style_bold=True)
        #
        # Coord([3, 3], [3, last_col], border=bc.get_border_thin(top=bc.MEDIUM)).draw(ws=ws, title="Дисциплина",
        #                                                                                  font_style_bold=True,
        #                                                                                  cell_len=Coord.CELL_LEN_NONE)
        for title_subject in self.subjects:
            self.last_col += 1
            Coord([4, self.last_col], border=bc.get_border_thin()).draw(ws=ws, title=title_subject,
                                                                        font_size=Coord.FONT_SIZE_MINI,
                                                                        font_style_bold=True,
                                                                        cell_len=Coord.CELL_LEN_AUTO_VERTICAL)

        Coord([3, 3], [3, self.last_col], border=bc.get_border_thin(top=bc.MEDIUM)).draw(ws=ws, title="Дисциплина",
                                                                                         font_style_bold=True,
                                                                                         cell_len=Coord.CELL_LEN_NONE)

    def create_statistic(self, ws: Union[WriteOnlyWorksheet, Worksheet]):
        self.last_col += 1
        title_right = {"Сред. балл": Coord([3, self.last_col], [4, self.last_col],
                                           border=bc.get_border_medium(bottom=bc.THIN))}
        self.draw_all(ws, title_right, font_style_bold=True)

        self.last_col += 1
        right_title_top = {
            "Пропущено часов": Coord([3, self.last_col], [3, self.last_col + 2],
                                     border=bc.get_border_medium(bottom=bc.THIN))}
        self.draw_all(ws, right_title_top, font_style_bold=True)

        right_title_bottom = {
            "Всего": Coord([4, self.last_col], border=bc.get_border_thin(left=bc.MEDIUM)),
            "По уваж.прич.": Coord([4, self.last_col + 1], border=bc.get_border_thin()),
            "По не уваж.прич.": Coord([4, self.last_col + 2], border=bc.get_border_thin(right=bc.MEDIUM)),
        }
        self.draw_all(ws, right_title_bottom)
        self.last_col += 2

    def draw_all(self, ws, d: dict, font=None, font_size=None, font_style_bold: bool = None):
        for title, coord in d.items():
            coord.draw(ws=ws, title=title, font=font, font_size=font_size, font_style_bold=font_style_bold)
