import os

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import pickle
import fake_useragent
from auth_data import kkmt_password, kkmt_login


class DriverHelper:
    URL_LOGIN = "https://ies.unitech-mo.ru/auth"
    HEADERS = {
        'user-agent': fake_useragent.UserAgent().random,
        'accept': '*/*'
    }

    def __init__(self):
        self.driver = self.get_driver()
        self.path = f'{os.getcwd()}\\parsing\\cookies\\'
        print(os.getcwd())
        self.cookie_file = f"{kkmt_login}_cookies"

    def get_driver(self):
        s = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=s)
        return driver

    def install_auth(self):
        print('install_auth')
        try:
            url = self.URL_LOGIN
            self.driver.get(url=url)
            print("Вход: Прогрузка страницы...")
            time.sleep(2)

            self.driver.find_element(By.XPATH, "//li[@class='user_session_data']").click()

            login = self.driver.find_element(By.XPATH, "//input[@name='login']")
            login.clear()
            login.send_keys(kkmt_login)

            password = self.driver.find_element(By.XPATH, "//input[@name='pass']")
            password.clear()
            password.send_keys(kkmt_password)

            self.driver.find_element(By.XPATH, "//a[@id='main_login_link']").click()

            time.sleep(1)
            #  cookies
            full_path = os.path.join(self.path + self.cookie_file)
            pickle.dump(self.driver.get_cookies(), open(full_path, "wb"))
        except Exception as ex:
            print(ex)

    def auth(self):
        try:
            self.driver.get(url=self.URL_LOGIN)
            print("Вход: Прогрузка страницы...")
            time.sleep(1)

            for cookie in pickle.load(open(self.cookie_file, "rb")):
                self.driver.add_cookie(cookie)
            self.driver.refresh()

        except Exception as ex:
            print(ex)

    def close(self):
        self.driver.close()
        self.driver.quit()
