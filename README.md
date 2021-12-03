# ParserSchedule

Данная программа парсит данные с сайта **https://ies.unitech-mo.ru** и создает отчет на основе поступившиз данных.   
Пример отчета: **result.xlsx  **

## Работа с кодом
Для начала надо скачать библиотеки. Для этого введите эту команду:
~~~~
pip install -r requirements.txt
~~~~
Также скачайте **Google Chrome**  

Далее вам надо перейти в файл **main.py** и в переменные **month**, **year**, **term**, ввести ваши данные.  
После этого перейдите в файл **auth_data.py** и введите логин и пароль от аккаунта.  
Если у вас нет аккаунта, можете протестировать код уже на спарсенных данных.  
Для этого закоментируйте код парсинга: с **10** по **11** строку и с **20** по **42** строку  
И раскоментируйте код с **44** по **48** строку.  

### Во время работы кода откроется браузер. НЕ ЗАКРЫВАЙ ЕГО!!!
  
После того, как код выполнит своё предназначение, откроется файл с именнем "Отчёт _название_учебной_группы_ _дата_отчёта_.xlsx"
  
На этом можно закончить
