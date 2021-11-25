import time
from bs4 import BeautifulSoup
from Url import Url


class Lessons:
    URL_LESSONS = "https://ies.unitech-mo.ru/studentplan"
    URL_MAIN_PAGE = "https://ies.unitech-mo.ru"

    def __init__(self, driver, term: int = -1):
        self.driver = driver
        self.term = term
        self.url = self.get_url()

    def get_url(self):
        params_url = {}
        if self.term != -1:
            params_url["sem"] = self.term
        return Url.get_url(self.URL_LESSONS, params_url)

    def parse(self):
        print("Получение ссылок на предметы...")
        url_lessons = self.get_url_lessons(
            Url.get_html(self.driver, self.url)
        )
        print(url_lessons)
        print(len(url_lessons))

    def get_url_lessons(self, html):
        soup = BeautifulSoup(html, 'html.parser')
        rows = soup.find('table', class_="student_plan_table").find_all('tr')
        url_lessons = dict()

        for row in rows[1:]:
            cols = row.find_all('td')
            lesson = str(cols[1])[4:-5]
            url = cols[-1].find('a').get('href')
            url_lessons[lesson] = self.URL_MAIN_PAGE + url
        return url_lessons




