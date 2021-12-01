from openpyxl import Workbook

from Schedule import Schedule
from BorderCoord import BorderCoord as bc
from FontCoord import FontCoord as fc
from AlignmentCoord import AlignmentCoord as ac
from Coord import Coord


class SkipSheet:

    def __init__(self, wb: Workbook, subjects: dict, schedule: dict, month: int, year: int, all_days: list,
                 name_sheet: str = None):
        if name_sheet is None:
            name_sheet = "(ЭН)"
        self.name_sheet = name_sheet
        self.wb = wb
        self.subjects = subjects
        self.schedule = schedule
        self.month = month
        self.year = year
        self.all_days = all_days

        self.ws = None
        self.col_end = None

        self.create_sheet()

    def create_sheet(self):
        self.ws = self.wb.create_sheet(self.name_sheet)

        self.create_students(start_row=1, start_col=1)
        self.create_skip(start_row=1, start_col=3)

    def create_students(self, start_row, start_col):
        """Создание записи со студентами"""
        row = start_row + 2
        col = start_col
        # Title
        Coord([start_row, col], [row, col], border=bc.get_border_thin(top=bc.BOLD, left=bc.BOLD), font=fc(bold=True)) \
            .draw(ws=self.ws, title="№ п/п")
        col += 1

        Coord([start_row, col], [row, col], border=bc.get_border_thin(top=bc.BOLD), font=fc(bold=True)) \
            .draw(ws=self.ws, title="ФИО")

        # Main
        row += 1
        students = list(list(self.subjects.values())[0].keys())
        for i in range(len(students)):
            Coord([row, start_col], border=bc.get_border_thin(left=bc.BOLD, bottom=bc.NONE)).draw(ws=self.ws,
                                                                                                  title=str(i + 1))
            Coord([row, col], border=bc.get_border_thin(bottom=bc.NONE)) \
                .draw(ws=self.ws, title=students[i], align=ac(horizontal=ac.LEFT), type_size=Coord.CELL_SIZE_TYPE_MAX)
            row += 1
        Coord([row, start_col], ws=self.ws, border=bc.get_border(top=bc.BOLD),
              type_size=Coord.CELL_SIZE_TYPE_MAX).draw()
        Coord([row, col], ws=self.ws, border=bc.get_border(top=bc.BOLD), type_size=Coord.CELL_SIZE_TYPE_MAX).draw()

    def create_skip(self, start_row, start_col):
        """Создание записи с парами и пропусками"""

        row = start_row
        col = start_col
        self.create_skip_title(row, col)

        row = start_row + 3
        self.create_skip_skip(row, col)

    def create_skip_title(self, start_row, start_col):
        """Создание титульной записи с датами и предметами"""
        row = start_row
        col = start_col

        count_days = 0
        for week in self.schedule:
            for day_week in range(1, 8):
                day = week.get(day_week, -1)
                if day != -1:
                    update_day = self.get_lessons(day)

                    border_subjects = bc.get_border(top=bc.BOLD, bottom=bc.MEDIUM)
                    border_number_lesson = bc.get_border_thin()

                    lessons = sorted(update_day.keys())
                    last_lesson = lessons[-1]
                    for lesson in lessons:
                        if lesson == last_lesson:
                            border_subjects = bc.get_border_medium(top=bc.BOLD, left=bc.NONE)
                            border_number_lesson = bc.get_border_thin(right=bc.MEDIUM)
                        # Название пары
                        Coord([row, col], ws=self.ws, title=update_day[lesson], border=border_subjects) \
                            .draw(font=fc(bold=True, size=fc.FONT_SIZE_MINI))
                        # Номер пары
                        Coord([row + 2, col], ws=self.ws, title=self.get_char_lesson(lesson), font=fc(bold=True)) \
                            .draw(cell_len=8, border=border_number_lesson)
                        col += 1
                    # Дата
                    Coord([row + 1, col - len(lessons)], [row + 1, col - 1], ws=self.ws, font=fc(bold=True), cell_len=8,
                          border=bc.get_border(left=bc.MEDIUM, right=bc.MEDIUM), title=self.all_days[count_days]).draw()
                    count_days += 1
        self.col_end = col - 1

    def get_lessons(self, day: dict):
        """Вернет словарь с правильной структурой пар"""
        number_lessons = len(day)  # Количество уроков
        if number_lessons >= 5:
            return day

        lessons = day.copy()
        min_lessons = min(lessons)
        max_lessons = max(lessons)
        left_lessons = 5 - number_lessons

        # 1,3,5 -> 1,2,3,4,5
        for num in range(min_lessons + 1, max_lessons):
            if lessons.get(num, None) is None:
                lessons[num] = dict()
                left_lessons -= 1
            if left_lessons == 0:
                break

        # 2,3 -> 1,2,3,4,5
        if left_lessons > 0:
            maxx = max(max_lessons, 5)
            for num in range(1, maxx + 1):
                if lessons.get(num, None) is None:
                    lessons[num] = dict()
                    left_lessons -= 1
                if left_lessons == 0:
                    break
        return lessons

    def get_char_lesson(self, lesson: int):
        if (lesson < 1) or (lesson > 8):
            return None
        lesson_chars = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"]
        return lesson_chars[lesson - 1]

    def create_skip_skip(self, start_row, start_col):
        row = start_row
        col = start_col
        row_lesson = start_row - 3
        row_date = start_row - 2
        row_num_lesson = start_row - 1
        col_student = start_col - 1

        name_students: list = sorted(list(list(self.subjects.values())[0].keys()))
        for i in range(start_row, len(name_students) + start_row):
            for j in range(start_col, self.col_end + 1):
                Coord([i, j], ws=self.ws, border=bc.get_border_thin()).draw()

        for title_lesson, students in self.subjects.items():
            for name, student_subjects in students.items():
                row_student = name_students.index(name) + start_row
                # print(Coord.get_title(ws=self.ws, row=row_student, column=col_student))
                for date, result in student_subjects.items():
                    current_date = Coord.get_title(ws=self.ws, row=row_date, column=col)
                    for col_day in range(col, self.col_end + 1):
                        intermediate_date = Coord.get_title(ws=self.ws, row=row_date, column=col_day)
                        if intermediate_date is None:
                            intermediate_date = current_date
                        if date == intermediate_date and \
                                title_lesson == Coord.get_title(ws=self.ws, row=row_lesson, column=col_day) and \
                                result == "Н":
                            Coord([row_student, col_day], ws=self.ws, title=result, border=bc.get_border_thin())\
                                .draw(type_size=Coord.CELL_SIZE_TYPE_MAX)
