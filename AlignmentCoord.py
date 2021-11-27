from openpyxl.styles import Alignment


class AlignmentCoord:
    """Работа с выравниванием текста"""

    CENTER = 'center'
    LEFT = 'left'
    RIGHT = 'right'

    def __init__(self, wrap_text: bool = True, horizontal: str = None, vertical: str = None):
        self.wrap_text = wrap_text
        if horizontal is None:
            horizontal = AlignmentCoord.CENTER
        self.horizontal = horizontal

        if vertical is None:
            vertical = AlignmentCoord.CENTER
        self.vertical = vertical

    def get(self):
        return Alignment(wrap_text=self.wrap_text, horizontal=self.horizontal, vertical=self.vertical)

