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
    year = 2021
    # # Получение всего расписания
    # schedule = Schedule(driver=driver.driver, month=month).parse()
    # print(schedule)
    #
    # Получение журнала успеваемости по всем предметам
    # subjects = Subjects(driver=driver.driver, term=7, month=month, year=year)
    # subjects.parse()
    # all_days = subjects.all_days

    # Получения данных о группе
    # user = User(driver.driver)
    # speciality = user.get_corporate_data(user.TEXT_SPECIALITY)
    # print(speciality)
    # group = user.get_corporate_data(user.TEXT_GROUP)
    # print(group)

    schedule = data.schedule
    subjects = data.subjects
    all_days = data.all_days
    group = "П1-18"
    speciality = "Программирование в компьютерных системах"

    # for week in schedule:
    #     print("----------------------------")
    #     for day in week:
    #         print(day)
    #         for key, value in week[day].items():
    #             print(f"{key}: {value}")
    #     print("----------------------------")
    # print("Парсинг расписания завершен успешно!")
    print(schedule)
    print("\n\n")
    print(subjects)
    print("\n\n")

    saver = Saver(schedule=schedule, subjects=subjects, group=group, speciality=speciality, all_days=all_days,
                  month=month, year=year)


if __name__ == '__main__':
    main()
