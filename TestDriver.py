import os

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


class TestDriver:

    def __init__(self):
        self.driver = self.get_driver()

    def get_driver(self):
        s = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=s)
        return driver

    def close(self):
        self.driver.close()
        self.driver.quit()
