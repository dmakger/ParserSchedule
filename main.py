import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup
import time
import pickle

from Schedule import Schedule
from Driver import DriverHelper


def main():
    driver = DriverHelper()
    # driver.install_auth()
    driver.auth()

    schedule = Schedule(driver.driver)
    schedule.parse()


if __name__ == '__main__':
    main()
