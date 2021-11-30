from Driver import DriverHelper
from Schedule import Schedule
from Subjects import Subjects
from User import User
from Saver import Saver
import data


def main():
    # driver = DriverHelper()
    # driver.install_auth()
    # # driver.auth()

    month = 10
    # # Получение всего расписания
    # schedule = Schedule(driver=driver.driver, month=month).parse()
    # print(schedule)
    #
    # # Получение журнала успеваемости по всем предметам
    # subjects = Subjects(driver=driver.driver, term=7, month=month).parse()
    # print(data.subjects)
    #
    # user = User(driver.driver)
    # speciality = user.get_corporate_data(user.TEXT_SPECIALITY)
    # print(speciality)
    # group = user.get_corporate_data(user.TEXT_GROUP)
    # print(group)

    schedule = data.schedule
    subjects = data.subjects
    group = "П1-18"
    speciality = "Программирование в компьютерных системах"

    saver = Saver(schedule=schedule, subjects=subjects, group=group, speciality=speciality, month=month)


if __name__ == '__main__':
    main()
