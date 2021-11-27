from typing import Union

from openpyxl import Workbook
from openpyxl.styles import borders
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet

from BorderCoord import BorderCoord as bc
from FontCoord import FontCoord as fc
from AlignmentCoord import AlignmentCoord as ac
from openpyxl.styles import Border


class Coord:
    CELL_LEN_NONE = 0
    CELL_LEN_AUTO_VERTICAL = 1
    CELL_LEN_AUTO_HORIZONTAL = 2

    CELL_WIDTH = 8.43
    CELL_HEIGHT = 15.75

    CELL_SIZE_TYPE_NONE = 0
    CELL_SIZE_TYPE_MAX = 1

    def __init__(self, start: list, end: list = None, font: fc = None, align: ac = None, border: Border = None,
                 title: str = "", width: int = CELL_WIDTH, height: int = CELL_HEIGHT,
                 cell_len: int = CELL_LEN_AUTO_HORIZONTAL, type_size: int = CELL_SIZE_TYPE_NONE,
                 ws: Union[WriteOnlyWorksheet, Worksheet] = None):
        self.start = start
        self.start_row: int = start[0]
        self.start_col: int = start[1]

        if end is None:
            end = start
        self.end = end
        self.end_row: int = end[0]
        self.end_col: int = end[1]

        if font is None:
            font = fc()
        self.font = font

        if align is None:
            align = ac()
        self.align = align

        if border is None:
            border = bc.get_border()
        elif type(border) == dict:
            border = bc.get_border(left=border.get('left', bc.NONE),
                                   right=border.get('right', bc.NONE),
                                   top=border.get('top', bc.NONE),
                                   bottom=border.get('bottom', bc.NONE))
        self.border = border

        if ws is None:
            ws = None
        self.ws = ws

        self.title = title
        self.cell_len = cell_len
        self.width = width
        self.height = height
        self.type_size = type_size

    def draw(self, ws: Union[WriteOnlyWorksheet, Worksheet] = None, title: str = None, font: fc = None,
             align: ac = None, type_size: int = None, cell_len: int = None, border: Border = None):
        """Отрисовка текста"""
        if ws is None:
            ws = self.ws
        if title is None:
            title = self.title
        if font is None:
            font = self.font
        if align is None:
            align = self.align
        if type_size is None:
            type_size = self.type_size
        if cell_len is None:
            cell_len = self.cell_len
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
        self.set_title(ws=ws, title=title, cell=cell)
        cell.font = font.get()
        cell.alignment = align.get()
        self.set_border(ws, border)

        self.set_cell_size(ws=ws, title=title, cell=cell, cell_len=cell_len, type_size=type_size)

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
            return self.CELL_WIDTH
        elif cell_len == self.CELL_LEN_AUTO_HORIZONTAL:
            cell_len_digit = (len(title) + 2) * 1.2
            # if self.start_row == self.end_row:
            return cell_len_digit

        elif cell_len == self.CELL_LEN_AUTO_VERTICAL:
            return (max([len(word) for word in title.split(' ')]) + 2) * 1.2
        else:
            return cell_len

    def get_size_row(self, width_column: int, title: str = None, ws: Union[WriteOnlyWorksheet, Worksheet] = None):
        if title is None:
            title = self.title
        if ws is None:
            ws: Union[WriteOnlyWorksheet, Worksheet] = self.ws

        current_row = ws.row_dimensions[self.start_row].height
        height = int(round(len(title) / width_column) * 15.50) + 10
        if current_row is None:
            current_row = 15.50
        if current_row + 10 > height:
            return current_row
        return height

    def set_cell_size(self, cell, cell_len, ws=None, title=None, width_column: int = None,
                      type_size: int = CELL_SIZE_TYPE_NONE):
        if title is None:
            title = self.title
        if ws is None:
            ws: Union[WriteOnlyWorksheet, Worksheet] = self.ws
        if width_column is None:
            width_column = self.get_size_column(title, cell_len)

        height = self.get_size_row(ws=ws, title=title, width_column=width_column)

        if ws.column_dimensions[cell.column_letter].width is None:
            ws.column_dimensions[cell.column_letter].width = self.CELL_WIDTH
        if ws.row_dimensions[cell.row].height is None:
            ws.row_dimensions[cell.row].height = self.CELL_HEIGHT

        if type_size == self.CELL_SIZE_TYPE_MAX:
            if ws.column_dimensions[cell.column_letter].width < width_column:
                ws.column_dimensions[cell.column_letter].width = width_column
            else:
                width_column = ws.column_dimensions[cell.column_letter].width
            self.width = width_column

            if ws.row_dimensions[cell.row].height > height:
                ws.row_dimensions[cell.row].height = height
            else:
                height = ws.row_dimensions[cell.row].height
            self.height = height

        elif type_size == self.CELL_SIZE_TYPE_NONE:
            ws.column_dimensions[cell.column_letter].width = width_column
            ws.row_dimensions[cell.row].height = height

    def set_title(self, cell, ws: Union[WriteOnlyWorksheet, Worksheet] = None, title: str = None):
        if ws is None:
            ws = self.ws
        if title is None:
            title = self.title

        if title[0] == '=':
            cell_pos = f"{cell.column_letter}{cell.row}"
            ws.write_formula(cell_pos, title)
        else:
            cell.value = title

