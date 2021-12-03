class SheetHelper:

    @staticmethod
    def get_char_lesson(lesson: int):
        """Вернет римское начертания числа"""
        if (lesson < 1) or (lesson > 8):
            return None
        lesson_chars = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"]
        return lesson_chars[lesson - 1]

    @staticmethod
    def get_lessons(day: dict):
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