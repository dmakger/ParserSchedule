from typing import Union

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.styles import Border
from openpyxl.styles import Side
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet


class Coord:
    FONT_FAMILY = "Times New Roman"

    FONT_SIZE_BIG = 16
    FONT_SIZE_MIDDLE = 12
    FONT_SIZE_REGULAR = 11
    FONT_SIZE_MINI = 10

    CENTER = 'center'
    LEFT = 'left'
    RIGHT = 'right'

    def __init__(self, start: list, end: list = None, ):
        self.start = start
        self.start_row: int = start[0]
        self.start_col: int = start[1]

        if end is None:
            end = start
        self.end = end
        self.end_row: int = end[0]
        self.end_col: int = end[1]

        self.ws = None
        self.title = None
        self.font = self.FONT_FAMILY
        self.font_size = self.FONT_SIZE_MIDDLE
        self.font_style_bold = False
        self.style_horizontal = self.CENTER
        self.style_vertical = self.CENTER

    def draw(self, ws: Union[WriteOnlyWorksheet, Worksheet] = None, title=None, font=None, font_size=None, font_style_bold: bool = None):
        if ws is None:
            ws = self.ws
        if title is None:
            title = self.title
        if font is None:
            font = self.font
        if font_size is None:
            font_size = self.font_size
        if font_style_bold is None:
            font_style_bold = self.font_style_bold

        self.merge(ws)
        # ws.set_column(1, 0, 15)
        # ws.row_dimensions[self.start_row].height = self.get_height()
        cell = ws.cell(row=self.start_row, column=self.start_col)
        cell.value = title
        cell.font = Font(name=font, size=font_size, bold=font_style_bold)
        cell.alignment = Alignment(wrap_text=True, horizontal=self.style_horizontal, vertical=self.style_vertical)
        # print(ws.row_dimensions[self.start_row].height)

    def merge(self, ws: Union[WriteOnlyWorksheet, Worksheet] = None):
        if ws is None:
            ws = self.ws

        ws.merge_cells(
            start_row=self.start_row, start_column=self.start_col,
            end_row=self.end_row, end_column=self.end_col
        )
