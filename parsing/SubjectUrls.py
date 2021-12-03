import datetime
from bs4 import BeautifulSoup
from parsing.Url import Url


class SubjectUrls:
    URL_LESSONS = "https://ies.unitech-mo.ru/studentplan"
    URL_MAIN_PAGE = "https://ies.unitech-mo.ru"

    def __init__(self, driver, term: int = -1, month: int = None, year: int = None):
        """
        Парсинг страницы для получения ссылок на предметы
        driver - нужен для перемещения по ссылкам
        term - нужный семестр
        month - месяц по которому нужна успеваемость
        """
        self.driver = driver
        self.term = term
        if month is None:
            month = datetime.date.today().month
        self.month = month

        if year is None:
            year = datetime.date.today().year
        self.year = year
        self.url = self.get_url()

    def get_url(self):
        """
        url:str -> вернет url созданный относительно семестра
        """
        params_url = {}
        if self.term != -1:
            params_url["sem"] = self.term
        return Url.get_url(self.URL_LESSONS, params_url)

    def parse(self, lessons: list = None):
        """
        parse:dict -> вернет предмет к ссылке.
        lessons:list -> сортировщик. Если он есть, то вернет только предметы которые есть в этом списке
        """
        print("Получение ссылок на учебные предметы")
        html = Url.get_html(self.driver, self.url)
        soup = BeautifulSoup(html, 'html.parser')
        rows = soup.find('table', class_="student_plan_table").find_all('tr')
        urls_lessons = dict()

        for row in rows[1:]:
            cols = row.find_all('td')
            lesson = str(cols[1])[4:-5]
            url = cols[-1].find('a').get('href')
            urls_lessons[lesson] = self.URL_MAIN_PAGE + url

        # Если есть сортировщик lessons
        if lessons is not None:
            new_urls_lessons = dict()
            for lesson in lessons:
                if lesson in urls_lessons:
                    new_urls_lessons[lesson] = urls_lessons[lesson]
            urls_lessons = new_urls_lessons

        return urls_lessons

