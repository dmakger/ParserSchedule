from openpyxl.styles import Font


class FontCoord:
    """Работа со стилем текста"""
    FONT_FAMILY_TIMES_NEW_ROMAN = "Times New Roman"

    FONT_SIZE_BIG = 16
    FONT_SIZE_MIDDLE = 12
    FONT_SIZE_REGULAR = 11
    FONT_SIZE_MINI = 10

    def __init__(self, family: str = None, size: int = None, bold: bool = False):
        if family is None:
            family = self.FONT_FAMILY_TIMES_NEW_ROMAN
        self.family = family

        if size is None:
            size = self.FONT_SIZE_MIDDLE
        self.size = size

        self.bold = bold

    def get(self):
        return Font(name=self.family, size=self.size, bold=self.bold)
