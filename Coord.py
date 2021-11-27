from typing import Union

from openpyxl import Workbook
from openpyxl.styles import borders
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet

from BorderCoord import BorderCoord as bc
from openpyxl.styles import Border


class Coord:
    FONT_FAMILY = "Times New Roman"

    FONT_SIZE_BIG = 16
    FONT_SIZE_MIDDLE = 12
    FONT_SIZE_REGULAR = 11
    FONT_SIZE_MINI = 10

    CENTER = 'center'
    LEFT = 'left'
    RIGHT = 'right'

    def __init__(self, start: list, end: list = None, border=bc.get_border()):
        self.start = start
        self.start_row: int = start[0]
        self.start_col: int = start[1]

        if end is None:
            end = start
        self.end = end
        self.end_row: int = end[0]
        self.end_col: int = end[1]

        if border is None:
            border = bc.get_border()
        elif type(border) == dict:
            border = bc.get_border(left=border.get('left', bc.NONE),
                                   right=border.get('right', bc.NONE),
                                   top=border.get('top', bc.NONE),
                                   bottom=border.get('bottom', bc.NONE))
        self.border = border

        self.ws = None
        self.title = None
        self.font = self.FONT_FAMILY
        self.font_size = self.FONT_SIZE_MIDDLE
        self.font_style_bold = False
        self.style_horizontal = self.CENTER
        self.style_vertical = self.CENTER

    def draw(self, ws: Union[WriteOnlyWorksheet, Worksheet] = None, title=None, font=None, font_size=None,
             font_style_bold: bool = None, border=None):
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
        if border is None:
            border = self.border
        elif type(border) == dict:
            border = bc.get_border(left=border.get('left', bc.NONE),
                                   right=border.get('right', bc.NONE),
                                   top=border.get('top', bc.NONE),
                                   bottom=border.get('bottom', bc.NONE))

        ws.merge_cells(
            start_row=self.start_row, start_column=self.start_col,
            end_row=self.end_row, end_column=self.end_col
        )
        # ws.set_column(1, 0, 15)
        # ws.row_dimensions[self.start_row].height = self.get_height()
        cell = ws.cell(row=self.start_row, column=self.start_col)
        cell.value = title
        cell.font = Font(name=font, size=font_size, bold=font_style_bold)
        cell.alignment = Alignment(wrap_text=True, horizontal=self.style_horizontal, vertical=self.style_vertical)
        # cell.border = border

        self.set_border(ws)
        # for i in range(self.start_row, self.end_row + 1):
        #     row_heights = [ws.row_dimensions[i].height for i in range(ws.max_row)]
        #     row_heights = [15 if rh is None else rh for rh in row_heights]
        ws.column_dimensions[cell.column_letter].width = (self.get_max_len_title(title) + 3)
        # ws.row_dimensions[self.start_row].height = 30
        # print(ws.row_dimensions[self.start_row].height)

    def set_border(self, ws: Union[WriteOnlyWorksheet, Worksheet] = None, border: Border = None):
        if ws is None:
            ws = self.ws
        if border is None:
            border = self.border
        elif type(border) == dict:
            border = bc.get_border(left=border.get('left', bc.NONE),
                                   right=border.get('right', bc.NONE),
                                   top=border.get('top', bc.NONE),
                                   bottom=border.get('bottom', bc.NONE))

        for i in range(self.start_row, self.end_row + 1):
            for j in range(self.start_col, self.end_col + 1):
                ws.cell(row=i, column=j).border = border

    def get_max_len_title(self, title: str = None):
        if title is None:
            title = self.title

        return max([len(word) for word in title.split(' ')])
