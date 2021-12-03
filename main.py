from Driver import DriverHelper
from Schedule import Schedule
from SubjectUrls import SubjectUrls
from Subjects import Subjects
from User import User
from Saver import Saver
import data


def main():
    driver = DriverHelper()
    driver.install_auth()
    # driver.auth()

    month = 10
    year = 2021
    term = 7
    # Получение всего расписания
    schedule = Schedule(driver=driver.driver, month=month, year=year)
    schedule_result = schedule.parse()
    all_lesson = schedule.all_lesson
    print(all_lesson)
    print()

    subject_urls = SubjectUrls(driver=driver.driver, term=term, month=month, year=year).parse(lessons=all_lesson)
    print(subject_urls)
    print()

    # Получение журнала успеваемости по всем предметам
    subjects = Subjects(driver=driver.driver, term=term, month=month, year=year, urls=subject_urls)
    subjects_result = subjects.parse()
    all_days = subjects.all_days
    print(subjects_result)
    print()

    # Получения данных о группе
    user = User(driver.driver)
    speciality = user.get_corporate_data(user.TEXT_SPECIALITY)
    print(speciality)
    group = user.get_corporate_data(user.TEXT_GROUP)
    print(group)
    print()
    driver.close()

    # schedule = data.schedule
    # subjects = data.subjects2
    # all_days = data.all_days
    # group = "П1-18"
    # speciality = "Программирование в компьютерных системах"

    # for week in schedule:
    #     print("----------------------------")
    #     for day in week:
    #         print(day)
    #         for key, value in week[day].items():
    #             print(f"{key}: {value}")
    #     print("----------------------------")
    # print("Парсинг расписания завершен успешно!")
    # print(schedule)
    # print("\n\n")
    # print(subjects)
    # print("\n\n")

    saver = Saver(schedule=schedule_result, subjects=subjects_result, group=group, speciality=speciality, all_days=all_days,
                  month=month, year=year).save()


if __name__ == '__main__':
    main()
