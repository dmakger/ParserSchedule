from typing import Union

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet

from BorderCoord import BorderCoord as bc
from FontCoord import FontCoord as fc
from AlignmentCoord import AlignmentCoord as ac
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
        self.create_statistic(ws)

    def create_students(self, ws: Union[WriteOnlyWorksheet, Worksheet]):
        """Создание записи со студентами"""
        # Title
        Coord([3, 1], [4, 1], border=bc.get_border_thin(top=bc.MEDIUM, left=bc.MEDIUM), font=fc(bold=True)).draw(ws=ws,
                                                                                                                 title="№ п/п")
        Coord([3, 2], [4, 2], border=bc.get_border_thin(top=bc.MEDIUM), font=fc(bold=True)).draw(ws=ws, title="ФИО")
        self.last_col = 2

        # Main
        students = list(list(self.subjects.values())[0].keys())
        current_row = 5
        for i in range(len(students)):
            Coord([current_row, 1], border=bc.get_border_thin(left=bc.MEDIUM)).draw(ws=ws, title=str(i + 1))
            Coord([current_row, 2], border=bc.get_border_thin()).draw(ws=ws, title=students[i],
                                                                      align=ac(horizontal=ac.LEFT),
                                                                      type_size=Coord.CELL_SIZE_TYPE_MAX)
            current_row += 1

        # Coord([current_row, 1], border=bc.get_border(top=bc.MEDIUM)).draw(ws=ws, type_size=Coord.CELL_SIZE_TYPE_MAX)
        # Coord([current_row, 2], border=bc.get_border(top=bc.MEDIUM)).draw(ws=ws, type_size=Coord.CELL_SIZE_TYPE_MAX)

    def create_subjects(self, ws: Union[WriteOnlyWorksheet, Worksheet]):
        """Создание записи с предметами и оценками"""
        text_mini = fc(size=fc.FONT_SIZE_MINI, bold=True)
        current_row = 4
        num_students = len(list(self.subjects.values())[0].keys())
        for title_subject in self.subjects:
            self.last_col += 1
            Coord([current_row, self.last_col], border=bc.get_border_thin()).draw(ws=ws, title=title_subject,
                                                                                  font=text_mini,
                                                                                  cell_len=Coord.CELL_LEN_AUTO_VERTICAL)
            for i in range(1, num_students + 1):
                Coord([current_row + i, self.last_col], border=bc.get_border_thin()).draw(ws=ws,
                                                                                          type_size=Coord.CELL_SIZE_TYPE_MAX)

        Coord([3, 3], [3, self.last_col], border=bc.get_border_thin(top=bc.MEDIUM)).draw(ws=ws, title="Дисциплина",
                                                                                         font=fc(bold=True),
                                                                                         cell_len=Coord.CELL_LEN_AUTO_VERTICAL)

        last_row = current_row + num_students + 1
        Coord([last_row, 1], [last_row, self.last_col], border=bc.get_border_medium()).draw(ws=ws,
                                                                                            title="ВСЕГО",
                                                                                            font=fc(bold=True),
                                                                                            align=ac(
                                                                                                horizontal=ac.RIGHT))

    def create_statistic(self, ws: Union[WriteOnlyWorksheet, Worksheet]):
        """Создание записи с успеваемостью"""
        self.last_col += 1
        current_row = 3
        Coord([current_row, self.last_col], [current_row + 1, self.last_col],
              border=bc.get_border_medium(bottom=bc.THIN), font=fc(bold=True), title="Сред. балл").draw(ws=ws)

        current_row += 2
        num_students = len(list(self.subjects.values())[0].keys())
        for i in range(1, num_students + 1):
            Coord([current_row, self.last_col], border=bc.get_border_thin(left=bc.MEDIUM, right=bc.MEDIUM), font=fc(),
                  title="").draw(ws=ws, type_size=Coord.CELL_SIZE_TYPE_MAX)
            current_row += 1

        self.last_col += 1
        Coord([3, self.last_col], [3, self.last_col + 2], border=bc.get_border_medium(bottom=bc.THIN),
              font=fc(bold=True), title="Пропущено часов").draw(ws=ws)

        Coord([4, self.last_col], border=bc.get_border_thin(left=bc.MEDIUM), title="Всего").draw(ws=ws)
        Coord([4, self.last_col + 1], border=bc.get_border_thin(), title="По уваж.прич.",
              cell_len=Coord.CELL_LEN_AUTO_VERTICAL).draw(ws=ws)
        Coord([4, self.last_col + 2], border=bc.get_border_thin(right=bc.MEDIUM), title="По не уваж.прич.",
              cell_len=Coord.CELL_LEN_AUTO_VERTICAL).draw(ws=ws)
        self.last_col += 2
