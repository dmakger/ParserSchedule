from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment

from Coord import Coord


class MainSheet:
    FONT_FAMILY = "Times New Roman"
    FONT_SIZE_BIG = 16
    FONT_SIZE_MIDDLE = 12
    FONT_SIZE_REGULAR = 11
    FONT_SIZE_MINI = 10

    def __init__(self, wb: Workbook, subjects: dict):
        self.subjects = subjects
        self.wb = wb
        self.create_sheet()

    def create_sheet(self):
        ws = self.wb.create_sheet("Ведомость усп.и посещ.")

        self.create_title(ws)

    def create_title(self, ws):
        nums_row_vertical = [3, 4]
        # title_left = {"№ п/п": 'A3:A4', "ФИО": 'B3:B4'}
        title_left = {
            "№ п/п": Coord([nums_row_vertical[0], 1], [nums_row_vertical[1], 1]),
            "ФИО": Coord([nums_row_vertical[0], 2], [nums_row_vertical[1], 2])
        }
        # ws.row_dimensions[nums_row_vertical[1]].height = 48
        self.draw(ws=ws, d=title_left, font_family=self.FONT_FAMILY, size=self.FONT_SIZE_MIDDLE, )

        title_bottom_middle = {}
        last_col = 2
        for title_subject in self.subjects:
            last_col += 1
            title_bottom_middle[title_subject] = Coord([nums_row_vertical[1], last_col])
        self.draw(ws=ws, d=title_bottom_middle, font_family=self.FONT_FAMILY, size=self.FONT_SIZE_MINI)

        # last_col += 1
        title_top_middle = {"Дисциплина": Coord([nums_row_vertical[0], 3], [nums_row_vertical[0], last_col])}
        self.draw(ws=ws, d=title_top_middle, font_family=self.FONT_FAMILY, size=self.FONT_SIZE_MIDDLE)

        last_col += 1
        title_right = {"Сред. балл": Coord([nums_row_vertical[0], last_col], [nums_row_vertical[1], last_col])}
        self.draw(ws=ws, d=title_right, font_family=self.FONT_FAMILY, size=self.FONT_SIZE_MIDDLE)

        last_col += 1
        right_title_top = {"Пропущено часов": Coord([nums_row_vertical[0], last_col], [nums_row_vertical[0], last_col + 3])}
        self.draw(ws=ws, d=right_title_top, font_family=self.FONT_FAMILY, size=self.FONT_SIZE_MIDDLE)

        right_title_bottom = {
            "Всего": Coord([nums_row_vertical[1], last_col]),
            "По уваж.прич.": Coord([nums_row_vertical[1], last_col + 1]),
            "По не уваж.прич.": Coord([nums_row_vertical[1], last_col + 2])
        }
        self.draw(ws=ws, d=right_title_bottom, font_family=self.FONT_FAMILY, size=self.FONT_SIZE_MIDDLE)
        last_col += 2

    def draw(self, ws, d: dict, font_family: str = "Calibri", size: int = 11):
        for title, coord in d.items():
            self.merge(ws, coord)
            cell = ws.cell(row=coord.start_row, column=coord.start_col)
            cell.value = title
            cell.font = Font(name=font_family, size=size)
            # cell.alignment.wrap_text = True
            cell.alignment = Alignment(wrap_text=True)

    def merge(self, ws, coord):
        ws.merge_cells(
            start_row=coord.start_row, start_column=coord.start_col,
            end_row=coord.end_row, end_column=coord.end_col
        )
