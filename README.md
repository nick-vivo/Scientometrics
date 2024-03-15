# Краткая инстукция по использованию:

Таблица с файлами.

Название файла  | Содержание файла
----------------|----------------------
main.ipynb       | Демонастрация получения выводов из шаблонов.
module/TableHandler.py       | Класс для работы с образцом.
resource  | Ввод и вывод данных
resource/Selection  | Обрезанный ввод данных, нужен для класса TableHandler.

## Создание класса TableHandler:

### Пример:

```python
from module.TableHandler import TableHandler

handler = TableHandler(pathForFile, gradesStudents, testScope, studentsNamesHeader)
```

где

    1. pathForFile - путь к файлу .xlsx таблицы
    2. gradeStudents - список заголовков оценок студентов по тестам.
    3. testScope - список заголовков оценок тестов студентами.
    4. studentsNamesHeader - при желании можно не указывать, тогда каждому студенту присвоиться свой id. Если указать и он является правильным, то имена студентов скопируются в итоговый вывод.

### Условия:

1. Нужно чтобы списоки вопросов, содержались в таблице.
2. pathForFile был путём к файлу .xlsx.


## Основные функции вывода таблиц TableHandler:
```python

1.handler.export_TableGradesStudent(pathForExport)

2.handler.export_TableGradesTest(pathForExport)

3.handler.export_TableLtiLsi(pathForExport)

4.handler.export_TableConclusion(nameFileForExport, pathForExport, nameOriginalForExport(имя для оригинальной таблицы), exportOriginal: да или нет(true, false))
```

1. Вывод таблицы оценок студентов по тестам с расчётами, по умолчанию экспортирует в эту папку(не создаст новой папки).(файл должен быть назван .xlsx)
   
2. Вывод таблицы оценок тестов студентами с расчётами, по умолчанию экспортирует в эту папку(не создаст новой папки).(файл должен быть назван .xlsx)

3.  Вывод таблицы LSI, LTI с расчётами, по умолчанию экспортирует в эту папку(не создаст новой папки).(файл должен быть назван .xlsx)

4.  Вывод всех таблиц в одной, по умолчанию экспортирует в эту папку(создаст новую папку, если указать в pathForExport).(файл должен быть назван .xlsx)


## Основные функции вывода графиков TableHandler:
```python

1.handler.export_PngPieBRSO(
                          fileToExport: str = "Путь к файлу для вывода",
                          nameHeader = 'Имя диаграммы')

2.handler.export_PngPieOTS(
                         fileToExport: str = "Путь к файлу для вывода", 
                         nameHeader = 'Имя диаграммы')

3. handler.export_PngPopularityTests(
                        fileToExport: str = "popularityTests.png")

3.handler.export_PngMotivation(fileToExportEducationMotivation: str = "Путь к первому файлу",
                             fileToExportMotivationEducation: str = "Путь ко второму файлу",
                             

4.handler.export_PngBenefits(fileToExport: str = "benefits.png",
                           headerBenefitsQuestion: str = "17. Какими средствами обучения вы преимущественно пользовались?",
                           typesBenefitsForPng: List[str] = ['Электронные учебники', 'Рабочие тетради', 'Видеолекции', 'Печатные учебники'],
                           typesBenefitsInTable: List[str] = ['Электронными учебниками', 'Рабочими тетрадями', 'Видеолекциями', 'Печатными учебниками'])

5. handler.export_PngConslission(
                              pathForExport: str = "./", 
                              namesFiles: List[str] = ["BRSO.png", "OTS.png", "Benefits.png", "Popularity.png", "Motivation.png", "Education.png"],
                              headerBenefitsQuestionInDataTable: str = "17. Какими средствами обучения вы преимущественно пользовались?",
                              typesBenefitsForPng: List[str] = ['Электронные учебники', 'Рабочие тетради', 'Видеолекции', 'Печатные учебники'],
                              typesBenefitsInDataTable: List[str] = ['Электронными учебниками', 'Рабочими тетрадями', 'Видеолекциями', 'Печатными учебниками']
                              )
6. handler.export_PngConclussionWithoutBenefits(
                              pathForExport: str = "./", 
                              namesFiles: List[str] = ["BRSO.png", "OTS.png", "Popularity.png", "Motivation.png", "Education.png"],
                              )
```

1. Вывод пирожковой диаграммы оценок студентов, округлённых по 0.5(параметр можно задать с помощью CONST_ROUND в классе или вне его)
   
2. Вывод пирожковой диаграммы оценок тестов студентами, округлённых по 0.5(параметр можно задать с помощью CONST_ROUND в классе или вне его)

3.  Вывод популярности заданий

4.  Вывод диаграммы корреляции между Успеваемостью и Мотивацией.
5.  Нужно знать как звучит вопрос в таблице, а также как написаны слова в таблице для поиска и подсчёта.
6.  Большинство параметров задано по умолчанию, при желании можно изменить.
7.  Вывод всех диаграмм без пособий, так как с пособиями есть проверка на вводимость данных typesBenefitys...



#     Конвертация и округление:

    Переменная GRADE_CONVERTED = 5, - в какую оценку идёт конвертация данных Average in X
    
    Переменная для округления оценок:
    ROUND_FACTOR = 0.5