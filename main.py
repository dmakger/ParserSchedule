from parsing.Driver import DriverHelper
from parsing.Schedule import Schedule
from parsing.SubjectUrls import SubjectUrls
from parsing.Subjects import Subjects
from parsing.User import User
from excel.Saver import Saver


def main():
    driver = DriverHelper()
    driver.install_auth()
    # driver.auth()

    # Месяц
    month = 11
    # Год
    year = 2021
    # Семестр
    term = 7
    # Получение всего расписания
    schedule = Schedule(driver=driver.driver, month=month, year=year)
    schedule_result = schedule.parse()
    all_lesson = schedule.all_lesson
    print(schedule_result)
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
    group = user.get_corporate_data(user.TEXT_GROUP)
    driver.close()

    # schedule = data.schedule
    # subjects = data.subjects2
    # all_days = data.all_days
    # group = "П1-18"
    # speciality = "Программирование в компьютерных системах"

    # print(schedule)
    # print("\n\n")
    # print(subjects)
    # print("\n\n")

    saver = Saver(schedule=schedule_result, subjects=subjects_result, group=group, speciality=speciality, all_days=all_days,
                  month=month, year=year).save()


if __name__ == '__main__':
    main()
