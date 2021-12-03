from bs4 import BeautifulSoup
from parsing.Url import Url


class User:
    URL_USER = "https://ies.unitech-mo.ru/user"
    TEXT_SPECIALITY = "Специальность / направление подготовки:"
    TEXT_GROUP = "Группа:"

    def __init__(self, driver):
        self.driver = driver
        self.soup = BeautifulSoup(Url.get_html(driver=self.driver, url=self.URL_USER), 'html.parser')

    def get_corporate_data(self, param: str = None):
        data = self.soup.find_all('div', class_="userpage_block_wrap")[1].find_all('div', class_="info")
        if param is None:
            corporate_data = list()
            for info in data:
                title = info.find("strong")
                if title:
                    corporate_data.append(info.get_text(strip=True).split(title.get_text(strip=True))[1])
            return corporate_data

        for info in data:
            title = info.find("strong")
            if title and (title.get_text(strip=True) == param):
                return info.get_text(strip=True).split(title.get_text(strip=True))[1]
        return None

