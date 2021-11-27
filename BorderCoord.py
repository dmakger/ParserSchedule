from openpyxl import Workbook
from openpyxl.styles import Border
from openpyxl.styles import Side


class BorderCoord:
    NONE = None
    THIN = 'thin'
    MEDIUM = 'medium'

    @staticmethod
    def get_border(left=NONE,
                   right=NONE,
                   top=NONE,
                   bottom=NONE):
        return Border(left=Side(style=left),
                      right=Side(style=right),
                      top=Side(style=top),
                      bottom=Side(style=bottom))

    @staticmethod
    def get_border_thin(left=THIN,
                        right=THIN,
                        top=THIN,
                        bottom=THIN):
        return Border(left=Side(style=left),
                      right=Side(style=right),
                      top=Side(style=top),
                      bottom=Side(style=bottom))

    @staticmethod
    def get_border_medium(left=MEDIUM,
                          right=MEDIUM,
                          top=MEDIUM,
                          bottom=MEDIUM):
        return Border(left=Side(style=left),
                      right=Side(style=right),
                      top=Side(style=top),
                      bottom=Side(style=bottom))
