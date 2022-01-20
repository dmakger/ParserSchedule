# ParserSchedule

Данная программа парсит данные с сайта **https://ies.unitech-mo.ru** и создает отчет на основе поступивших данных.   
Пример отчета файл: **result.xlsx**  

## Требования
Для начала надо скачать библиотеки. Для этого введите эту команду:
~~~~
pip install -r requirements.txt
~~~~
Также скачайте **Google Chrome** ==> **https://www.google.ru/chrome/**  

## Работа с кодом  
Вам надо перейти в файл **main.py** и в переменные **month**, **year**, **term**, ввести ваши данные.  
**month** - месяц для отчёта  
**year** - год для отчёта  
**term** - семестр для отчёта  

### 1. У вас есть логин и пароль от сайта
После этого перейдите в файл **auth_data.py** и введите логин и пароль от своего аккаунта.  
Для этого на строке **26** у переменной **key** поменяйте значение на _REAL_  

### 2. Просто проверить создание отчёта
Если у вас нет аккаунта, можете протестировать код уже на спарсенных данных.  
Для этого на строке **26** у переменной **key** поменяйте значение на _TEST_  

### Во время работы кода откроется браузер. НЕ ЗАКРЫВАЙ ЕГО!!!  
После того, как код выполнит своё предназначение, откроется файл с именнем _"Отчёт название-учебной-группы дата-отчёта.xlsx"_  
Например: _"Отчёт П1-18 10.2021"_  
Файл будет сохранён в папку с программой
  
На этом можно закончить
