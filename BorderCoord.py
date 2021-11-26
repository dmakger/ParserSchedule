from openpyxl import Workbook
from openpyxl.styles import Border
from openpyxl.styles import Side


class BorderCoord:
    THIN = Border(left=Side(style='thin'),
                  right=Side(style='thin'),
                  top=Side(style='thin'),
                  bottom=Side(style='thin'))
    MEDIUM = Border(left=Side(style='medium'),
                    right=Side(style='medium'),
                    top=Side(style='medium'),
                    bottom=Side(style='medium'))

    # def c