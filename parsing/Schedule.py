import datetime
from bs4 import BeautifulSoup
from parsing.Url import Url


class Schedule:
    URL_SCHEDULE = "https://ies.unitech-mo.ru/schedule"

    @staticmethod
    def get_all_days(month: int, year: int):
        start_day = datetime.date(month=month, year=year, day=1)
        days = list()
        day = start_day
        while day.month == start_day.month:
            days.append(day.strftime('%d.%m.%Y'))
            day += datetime.timedelta(days=1)
        return days

    def __init__(self, driver, month: int = None, year: int = None):
        """
        driver - нужен для перемещения по ссылкам
        month - месяц по которому нужна успеваемость
        year - год по которому нужна успеваемость
        """
        self.driver = driver
        if month is None:
            month = datetime.date.today().month
        self.month = month

        if year is None:
            year = datetime.date.today().year
        self.year = year

        self.start_day = datetime.date(self.year, self.month, 1)
        self.end_day = self.get_end_day()

        # Все уроки
        self.all_lesson = None

    def get_end_day(self):
        """day -> номер последнего дня месяца"""
        end_date = self.start_day + datetime.timedelta(days=30)
        if end_date.month == self.month:
            return end_date
        end_date -= datetime.timedelta(days=1)
        if end_date.month == self.month:
            return end_date
        end_date -= datetime.timedelta(days=1)
        if end_date.month == self.month:
            return end_date
        end_date -= datetime.timedelta(days=1)
        if end_date.month == self.month:
            return end_date

    def get_range_date(self):
        """range_date -> список словарей с датой начала недели и концом недели"""
        range_date = []
        first_date = self.start_day - datetime.timedelta(days=self.start_day.weekday())
        end_date = first_date + datetime.timedelta(days=6)
        while end_date.month == self.month:
            range_date.append({
                "min": first_date.strftime('%d.%m.%Y'),
                "max": end_date.strftime('%d.%m.%Y')
            })
            first_date = end_date + datetime.timedelta(days=1)
            end_date = first_date + datetime.timedelta(days=6)
        if range_date[-1]['max'] != self.end_day.strftime('%d.%m.%Y'):
            range_date.append({
                "min": first_date.strftime('%d.%m.%Y'),
                "max": end_date.strftime('%d.%m.%Y')
            })

        return range_date

    def get_schedule(self, html, start: int = 1, end: int = 7):
        """
        schedule:dict -> вернет все пары в течении недели.
        А именно словарь с числом дня недели пары (где 1 - понедельник) и
        словарь с номерами пар к названию пары
        week - ограничивает день недели
        """
        soup = BeautifulSoup(html, 'html.parser')
        table = soup.find('table', class_='schedule_day_time_table')
        schedule = dict()

        if self.all_lesson is None:
            self.all_lesson = list()

        items = table.find('tbody').find_all('tr')
        # print(items)
        for item in items:
            cols = item.find_all('td')
            for td in cols:
                info = td.find('div', class_='time_table_item_validated')
                if info and not td.find('div', class_='item_holiday'):
                    index = int(td.get('data-stt-day'))
                    if (index >= start) and (index <= end):
                        lesson = info.get('data-content').split('.')[1].strip()
                        if schedule.get(index, -1) == -1:
                            schedule[index] = dict()
                        schedule[index][int(td.get('data-stt-time'))] = lesson
                        self.all_lesson.append(lesson)
        self.all_lesson = list(set(self.all_lesson))
        return schedule

    def parse(self):
        """
        data:list -> вернет все пары в течении месяца
        """
        range_date = self.get_range_date()
        print(range_date)
        pages_count = len(range_date)
        data = list()

        # Проверка на первую неделю
        print(f"Парсинг страницы расписания. 1 из {pages_count}...")
        html = Url.get_html(self.driver, self.URL_SCHEDULE, params={
            "d": range_date[0]["min"] + "+-+" + range_date[0]["max"]
        })
        data.append(self.get_schedule(html=html, start=self.start_day.weekday()+1))

        # Проверка на все недели кроме последней
        for i in range(1, pages_count-1):
            print(f"Парсинг страницы расписания. {i + 1} из {pages_count}...")
            html = Url.get_html(self.driver, self.URL_SCHEDULE, params={
                "d": range_date[i]["min"] + "+-+" + range_date[i]["max"]
            })
            data.append(self.get_schedule(html))

        # Проверка на последнею неделю
        print(f"Парсинг страницы расписания. {pages_count} из {pages_count}...")
        html = Url.get_html(self.driver, self.URL_SCHEDULE, params={
            "d": range_date[-1]["min"] + "+-+" + range_date[-1]["max"]
        })
        print(self.end_day.weekday())
        data.append(self.get_schedule(html=html, end=self.end_day.weekday()+1))
        print("Парсинг расписания завершен успешно!")

        return data

