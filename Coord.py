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

    CELL_LEN_NONE = 0
    CELL_LEN_AUTO_VERTICAL = 1
    CELL_LEN_AUTO_HORIZONTAL = 2

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
        self.cell_len = self.CELL_LEN_AUTO_HORIZONTAL

    def draw(self, ws: Union[WriteOnlyWorksheet, Worksheet] = None, title=None, font=None, font_size=None,
             font_style_bold: bool = None, cell_len=None, style_horizontal=None,
             style_vertical=None, border=None):
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
        if cell_len is None:
            cell_len = self.cell_len
        if style_horizontal is None:
            style_horizontal = self.style_horizontal
        if style_vertical is None:
            style_vertical = self.style_vertical
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
        cell = ws.cell(row=self.start_row, column=self.start_col)
        cell.value = title
        cell.font = Font(name=font, size=font_size, bold=font_style_bold)
        cell.alignment = Alignment(wrap_text=True, horizontal=style_horizontal, vertical=style_vertical)
        self.set_border(ws, border)

        width_column = self.get_size_column(title=title, cell_len=cell_len)
        ws.column_dimensions[cell.column_letter].width = width_column
        ws.row_dimensions[cell.row].height = self.get_size_row(ws=ws, title=title, width_column=width_column)
        print(f"{title} = {width_column}x{self.get_size_row(ws=ws, title=title, width_column=width_column)} -> {cell.column_letter}:{cell.row}")

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

    def get_size_column(self, title, cell_len=CELL_LEN_AUTO_HORIZONTAL):
        if title is None:
            title = self.title

        # return max([len(word) for word in title.split(' ')]) + 10
        if cell_len == self.CELL_LEN_NONE:
            return 8.09
        elif cell_len == self.CELL_LEN_AUTO_HORIZONTAL:
            return len(title) + 4
        elif cell_len == self.CELL_LEN_AUTO_VERTICAL:
            # print(title + ": VERTICAL")
            # [print(len(word)) for word in title.split(' ')]
            return max([len(word) for word in title.split(' ')]) + 10
        else:
            return cell_len

    def get_size_row(self, width_column: int, title: str = None, ws: Union[WriteOnlyWorksheet, Worksheet] = None):
        if title is None:
            title = self.title
        if ws is None:
            ws: Union[WriteOnlyWorksheet, Worksheet] = self.ws

        current_row = ws.row_dimensions[self.start_row].height
        if current_row is None:
            current_row = 15.50
        height = int(round(len(title) / width_column) * 15.50) + 10
        if current_row + 10 > height:
            return current_row
        return height
