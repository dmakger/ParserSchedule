import datetime
import time
from bs4 import BeautifulSoup
from Url import Url
from selenium.webdriver.common.by import By


class Schedule:
    URL_SCHEDULE = "https://ies.unitech-mo.ru/schedule"

    def __init__(self, driver, month: int = None, year: int = None):
        self.driver = driver
        self.date = datetime.date
        if month is None:
            self.month = self.date.today().month
        else:
            self.month = month

        if year is None:
            self.year = self.date.today().year
        else:
            self.year = year

        self.start_day = self.date(self.year, self.month, 1)
        self.end_day = self.get_end_day()

    def get_end_day(self):
        """day -> номер последнего дня месяца"""
        end_date = self.start_day + datetime.timedelta(days=28)
        if end_date.month == self.month:
            return end_date
        end_date += datetime.timedelta(days=1)
        if end_date.month == self.month:
            return end_date
        end_date += datetime.timedelta(days=1)
        if end_date.month == self.month:
            return end_date
        end_date += datetime.timedelta(days=1)
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
        range_date.append({
            "min": first_date.strftime('%d.%m.%Y'),
            "max": end_date.strftime('%d.%m.%Y')
        })

        return range_date

    def get_schedule(self, html):
        """
        schedule:dict -> вернет все пары в течении недели.
        А именно словарь с числом дня недели пары (где 1 - понедельник) и
        словарь с номерами пар к названию пары
        """
        soup = BeautifulSoup(html, 'html.parser')
        table = soup.find('table', class_='schedule_day_time_table')
        num_days = len(table.find('thead').find_all('th')) - 1
        # schedule = [dict() for i in range(num_days)]
        schedule = dict()

        items = table.find('tbody').find_all('tr')
        for item in items:
            cols = item.find_all('td')
            for td in cols:
                info = td.find('div', class_='time_table_item_validated')
                if info and not td.find('div', class_='item_holiday'):
                    index = td.get('data-stt-day')
                    lesson = info.get('data-original-title').split('(')[-1].split(' - ')[1]
                    if schedule.get(index, -1) == -1:
                        schedule[index] = dict()
                    schedule[index][td.get('data-stt-time')] = lesson
        return schedule

    def parse(self):
        """
        data:list -> вернет все пары в течении месяца
        """
        range_date = self.get_range_date()
        pages_count = len(range_date)
        data = list()

        # for i in range(pages_count):
        for i in range(1):
            print(f"Парсинг страницы {i + 1} из {pages_count}...")
            html = Url.get_html(self.driver, self.URL_SCHEDULE, params={
                "d": range_date[i]["min"] + "+-+" + range_date[i]["max"]
            })
            data.append(self.get_schedule(html))

        for week in data:
            print("----------------------------")
            for day in week:
                print(day)
                for key, value in week[day].items():
                    print(f"{key}: {value}")
            print("----------------------------")

        return data

