from parsing.Driver import DriverHelper
from parsing.Schedule import Schedule
from parsing.SubjectUrls import SubjectUrls
from parsing.Subjects import Subjects
from parsing.User import User
from excel.Saver import Saver
import data

REAL = 1
TEST = 2

START = True
STOP = False


def main():
    # Месяц
    month = 9
    # Год
    year = 2022
    # Семестр
    term = 1

    # REAL - парсинг
    # TEST - тестовые значения
    key = REAL

    # Запуск создания файла
    start = START

    if key == REAL:
        driver = DriverHelper()
        driver.install_auth()
        # driver.auth()

        # Получение всего расписания
        schedule = Schedule(driver=driver.driver, month=month, year=year)
        schedule_result = schedule.parse()
        print(schedule_result)
        all_lesson = schedule.all_lesson
        print(f'Предметы в течении месяца: {all_lesson}')
        print('Расписание предметов:')
        print(*schedule_result, sep='\n')
        # print(schedule_result)
        print()

        subject_urls = SubjectUrls(driver=driver.driver, term=term, month=month, year=year).parse(lessons=all_lesson)
        print(subject_urls)
        print()

        # Получение журнала успеваемости по всем предметам
        subjects = Subjects(driver=driver.driver, term=term, month=month, year=year, urls=subject_urls)
        subjects_result = subjects.parse()
        all_days = subjects.all_days
        print("---")
        print(f'Дни в которые были пары {all_days}')
        print("---")
        print('Результаты в течении месяца')
        print(subjects_result)
        print()

        # Получения данных о группе
        user = User(driver.driver)
        speciality = user.get_corporate_data(user.TEXT_SPECIALITY)
        group = user.get_corporate_data(user.TEXT_GROUP)
        driver.close()

    if key == TEST:
        schedule_result = data.schedule
        subjects_result = data.subjects
        all_days = data.all_days12
        group = "П1-18"
        speciality = "Программирование в компьютерных системах"

        print(schedule_result)
        print("\n\n")
        print(subjects_result)
        print("\n\n")

    if start == START:
        Saver(schedule=schedule_result, subjects=subjects_result, group=group, speciality=speciality,
              all_days=all_days, month=month, year=year).save()


if __name__ == '__main__':
    main()
