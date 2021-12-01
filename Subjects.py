import datetime
from bs4 import BeautifulSoup
from Url import Url


class Subjects:
    URL_LESSONS = "https://ies.unitech-mo.ru/studentplan"
    URL_MAIN_PAGE = "https://ies.unitech-mo.ru"

    def __init__(self, driver, term: int = -1, month: int = None, year: int = None):
        """
        driver - нужен для перемещения по ссылкам
        term - нужный семестр
        month - месяц по которому нужна успеваемость
        """
        self.driver = driver
        self.term = term
        self.date = datetime.date
        if month is None:
            month = self.date.today().month
        self.month = month

        if year is None:
            year = self.date.today().year
        self.year = year
        self.url = self.get_url()

        self.all_days = None

    def get_url(self):
        """
        url:str -> вернет url созданный относительно семестра
        """
        params_url = {}
        if self.term != -1:
            params_url["sem"] = self.term
        return Url.get_url(self.URL_LESSONS, params_url)

    def parse(self):
        """
        gradebook:dict -> вернет предметы к студенту к его успеваемости.
        """
        print("Получение ссылок на предметы...")
        url_subjects = self.get_url_lessons(Url.get_html(self.driver, self.url))
        count_subjects = len(url_subjects)
        count = 1
        gradebook = dict()
        self.all_days = list()
        for title, url in url_subjects.items():
            print(f"Парсинг страниц предметов. {count} из {count_subjects}...")
            count += 1
            html = Url.get_html(self.driver, url)
            gradebook[title] = self.get_info_subject(html)
        print("Парсинг предметов завершен успешно!")
        self.all_days.sort()
        return gradebook

    def get_url_lessons(self, html):
        """
        url_lessons:dict -> вернет предметы к студенту к его успеваемости.
        """
        soup = BeautifulSoup(html, 'html.parser')
        rows = soup.find('table', class_="student_plan_table").find_all('tr')
        url_lessons = dict()

        for row in rows[1:]:
            cols = row.find_all('td')
            lesson = str(cols[1])[4:-5]
            url = cols[-1].find('a').get('href')
            url_lessons[lesson] = self.URL_MAIN_PAGE + url
        return url_lessons

    def get_info_subject(self, html):
        """
        peoples:dict -> вернет успеваемость каждого студента
        """
        soup = BeautifulSoup(html, 'html.parser')
        table = soup.find('table', class_="scorestable")
        dates = self.get_header(table)

        peoples = dict()
        rows = table.find('tbody').find_all('tr')
        for row in rows:
            name = row.find('span', class_="j_filter_by_fio").get_text().split('. ')[1]
            peoples[name] = dict()
            cols = row.find_all('td', class_="journal_ltype_0")
            for i in range(len(cols)):
                col_text = cols[i].get_text()
                if (col_text != "") and (dates.get(i, -1) != -1):
                    peoples[name][dates[i]] = col_text
        return peoples

    def get_header(self, table):
        """
        dates:dict -> вернет все даты подходящие по месяцу
        """
        dates = dict()
        rows = table.find('thead').find_all('th')[1:]
        for i in range(len(rows)):
            date = rows[i].find('a')
            if date:
                date_text = date.get_text(strip=True)
                if (date_text[0].isdigit()) and (int(date_text.split('.')[1]) == self.month):
                    dates[i] = date_text
                    if date_text not in self.all_days:
                        self.all_days.append(date_text)
        return dates
