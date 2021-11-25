from Driver import DriverHelper
from Schedule import Schedule
from Lessons import Lessons


def main():
    driver = DriverHelper()
    # driver.install_auth()
    driver.auth()

    # Получение всего расписания
    # schedule = Schedule(driver=driver.driver, month=10).parse()

    lessons = Lessons(driver=driver.driver, term=7).parse()


if __name__ == '__main__':
    main()
