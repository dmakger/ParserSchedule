import datetime
from bs4 import BeautifulSoup
from Url import Url


class Subjects:

    def __init__(self, driver, urls: dict, term: int = -1, month: int = None, year: int = None):
        """
        driver - нужен для перемещения по ссылкам
        term - нужный семестр
        month - месяц по которому нужна успеваемость
        """
        self.driver = driver
        self.urls = urls
        self.term = term
        if month is None:
            month = datetime.date.today().month
        self.month = month

        if year is None:
            year = datetime.date.today().year
        self.year = year

        self.all_days = None

    def parse(self):
        """
        gradebook:dict -> вернет предметы к студенту к его успеваемости.
        """
        print("Получение ссылок на предметы...")
        count_subjects = len(self.urls)
        count = 1
        gradebook = dict()
        self.all_days = list()
        for title, url in self.urls.items():
            print(f"Парсинг страниц с учебными предметами: {count} из {count_subjects}...")
            count += 1
            html = Url.get_html(self.driver, url)
            gradebook[title] = self.get_info_subject(html)
        print("Парсинг учебных предметов завершен!")
        self.all_days.sort()
        return gradebook

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
