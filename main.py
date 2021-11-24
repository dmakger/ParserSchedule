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
from auth_data import kkmt_password, kkmt_login
from Driver import DriverHelper


def main():
    driver = DriverHelper()
    driver.auth()


if __name__ == '__main__':
    main()
