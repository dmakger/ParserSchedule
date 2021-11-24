import datetime
import time
from bs4 import BeautifulSoup
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

    def get_url(self, url: str, params: dict = None):
        """url:str -> отформатированный url"""
        if (params is None) or (len(params) == 0):
            return url
        else:
            params_str = ""
            for key, value in params.items():
                params_str += key + "=" + value + "&"
            return url + "?" + params_str[:-1]

    def get_html(self, url: str, params: dict = None):
        """html:str -> вернет странницу html"""
        self.driver.get(self.get_url(url, params))
        print("Прогрузка страницы...")
        time.sleep(2)
        return self.driver.page_source

    def get_schedule(self, html):
        soup = BeautifulSoup(html, 'html.parser')
        table = soup.find('table', class_='schedule_day_time_table')
        num_days = len(table.find('thead').find_all('th')) - 1
        schedule = [list() for i in range(num_days)]

        items = table.find('tbody').find_all('tr')
        for item in items:
            cols = item.find_all('td')
            for td in cols:
                info = td.find('div', class_='time_table_item_validated')
                if info and not td.find('div', class_='item_holiday'):
                    index = int(td.get('data-stt-day')) - 1
                    schedule[index].append(td.get('data-stt-time'))
        return schedule

    def parse(self):
        range_date = self.get_range_date()
        pages_count = len(range_date)
        data = list()

        for i in range(pages_count):
        # for i in range(1):
            print(f"Парсинг страницы {i + 1} из {pages_count}...")
            html = self.get_html(self.URL_SCHEDULE, params={
                "d": range_date[i]["min"] + "+-+" + range_date[i]["max"]
            })
            data.append(self.get_schedule(html))

        for week in data:
            print("----------------------------")
            for day in week:
                print(day)

