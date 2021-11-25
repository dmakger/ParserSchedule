from Driver import DriverHelper
from Schedule import Schedule
from Subjects import Subjects


def main():
    driver = DriverHelper()
    # driver.install_auth()
    driver.auth()

    # Получение всего расписания
    schedule = Schedule(driver=driver.driver, month=10).parse()

    # Получение журнала успеваемости по всем предметам
    subjects = Subjects(driver=driver.driver, term=7, month=10).parse()
    print(subjects)


if __name__ == '__main__':
    main()
