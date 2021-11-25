class Coord:

    def __init__(self, start: list, end: list = None):
        self.start = start
        self.start_row = start[0]
        self.start_col = start[1]

        if end is None:
            end = start
        self.end = end
        self.end_row = end[0]
        self.end_col = end[1]
