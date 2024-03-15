import os
import shutil
from typing import List

import openpyxl
import numpy as np
import pandas as pd

import matplotlib.pyplot as plt

from .Exceptions.BadTable import BadTable as BadTable
from .Exceptions.BadNameHeaders import BadNameHeaders as BadNameHeaders


class TableHandler:
    
    # В какую оценку всё конвертировать в пятибальную и тд
    GRADE_CONVERTED = 5
    #Как происходит округление оценок
    ROUND_FACTOR = 0.5

    @staticmethod
    def check_table(path_to_table: str, 
                    headers: List[str], 
                    headersGrades: List[str], 
                    headersTestScope: List[str]) -> bool:
        """
        Проверяет таблицу на наличие ошибок.

        Args:
            path_to_table: str Путь к таблице.
            headers: List[str] Заголовки таблицы.
            questions: List[str] Вопросы в таблице.
            assessment_issues: List[str] Оценки тестов в таблице.

        Raise:
            - 1: Количество вопросов не совпадает с количеством оценок тестов.
            - 2: Количество вопросов не может быть нулевым, проверь таблицу
            - 3: Предоставленный файл не является формата .xlsx.
            - 4: Вопросы или оценки тестов не находятся в заголовках.
            - 5: Заголовки таблицы не совпадают с заголовками header.
        Return:
            - True: Ошибок нет.
        """
        workbook = openpyxl.load_workbook(path_to_table, read_only=True)
        worksheet = workbook.active

        headersTable = [cell.value for cell in next(worksheet.iter_rows())]
        
        if len(headersGrades) != len(headersTestScope):
            raise BadTable("Количество вопросов не совпадает с количеством оценок тестов")
        elif len(headersGrades) == 0:
            raise BadTable("Количество вопросов не может быть нулевым, проверь таблицу")
        elif not all(x in headers for x in headersGrades):
            raise BadTable("Вопросы оценок за контрольные не находятся в заголовках")
        elif not all(x in headers for x in headersTestScope):
            raise BadTable("Вопросы оценки тестов не находятся в заголовках")
        elif not all(x in headersTable for x in headers):
                raise BadTable("Заголовки таблицы не совпадают с заголовками header")

        return True

    @staticmethod    
    def customRound(number:float, cRound: float = ROUND_FACTOR) -> int:
        """
        Округляет число с заданной точностью().

        Args:
            - number: float число для округления.
            - cRound: float глубина округления(по умолчанию 0.5).
        Returns:
            - Округлённое число.
        """
        if number - int(number) >= cRound: 
            return int(number) + 1
        else:
            return int(number)
        
    
    @staticmethod
    def calculateAverage(row):
        """
        Суммирует строку и если в ней есть нулевые значения, то не учитывает их
        в общей сумме для нахождения среднего арифметического

        Args:
            - row: строка из pd.DataFrame

        Return:
            - Среднее арифметическое строки, не учитывая нули.
        """
        non_zero_values = row[row != 0]
        if len(non_zero_values) > 0:
            return non_zero_values.sum() / len(non_zero_values)
        else:
            return 0
    
    
    
    def __init__(self, 
                 path_to_table: str,
                 headersGrades: List[str], 
                 headersTestScore: List[str], 
                 headerNamesStudents: str = ""):
        """
        Конструктор для обработки таблицы

        Args:
            path_to_table: str Путь к таблице.
            headersGrades: List[str] Оценки студентов по тестам
            headersTestScore: List[str] Оценки тестов студентами.

        Raise:
            - 1: Количество вопросов не совпадает с количеством оценок тестов.
            - 2: Количество вопросов не может быть нулевым, проверь таблицу
            - 3: Предоставленный файл не является формата .xlsx.
            - 4: Вопросы или оценки тестов не находятся в заголовках таблицы.
        """
        headers = headersGrades + headersTestScore
        
        TableHandler.check_table(path_to_table, headers, 
                                 headersGrades, headersTestScore)

        self.__headersGradesStudents = list(headersGrades)
        self.__headersTestScore = list(headersTestScore)
        self.__headerStudentsName = headerNamesStudents
        self.__data_table = pd.read_excel(path_to_table)

        try:
            self.__names = self.__data_table[headerNamesStudents].to_list()
        except KeyError:
            self.__names = []
            for i in range(1, self.__data_table.shape[0] + 1):
                self.__names.append(str(i))


    def createTableGradesStudents(self, 
                                  nameHeadersColumns_Sum_Average_Round:List[str] = ["Sum", f"Average {GRADE_CONVERTED}", f"Round {ROUND_FACTOR}"]
                                  ) -> pd.DataFrame:
        """
        Создаёт таблицу оценок студентов pd.DataFrame, добавляя к ней колонки Sum,
        Average, Round.
        - Sum - сумма всех чисел строки
        - Average X - округление числа в GRADE_CONVERTED оценку, находя при этом 
        среднее
        - Round X - округление числа с заданным ROUND_FACTOR. Параметр округляет 
        число с заданной точностью.
        Список список вопросов берётся из переменной self.__headersGradesStudents.

        Args:
            - nameHeadersColumns_Sum_Average_Round - список названия колонок. 
            При желании названия по умолчанию можно изменить.

        Return:
            - Таблица pd.DataFrame с новыми колонками.
        """
        return TableHandler.createTableWithNewColumns_SumAverageRound(self.__data_table, 
                                                                      self.__headersGradesStudents, 
                                                                      nameHeadersColumns_Sum_Average_Round)

    
    def createTableGradesStudentsToView(self, 
                                        nameColumnStuneds: str = "Students",
                                        nameHeadersColumn_Sum_Average_Round:List[str] = ["Sum", f"Average {GRADE_CONVERTED}", f"Round {ROUND_FACTOR}"],
                                        nameHeadersString_Max_Sum_Average:List[str] = ['Max', 'Sum', 'Average', f"Average {GRADE_CONVERTED}"]) -> pd.DataFrame:
        """
        Создаёт таблицу оценок студентов pd.DataFrame, добавляя к ней колонки Sum,
        Average, Round.
        - Sum - сумма всех чисел строки
        - Average X - округление числа в GRADE_CONVERTED оценку, находя при этом
        среднее
        - Round X - округление числа с заданным ROUND_FACTOR. Параметр округляет
        число с заданной точностью.
        Также добалвяются строки таблицы 'Max' 'Sum' 'Average' 'Average X' 
        - Max - максимальное число в колонке,
        - Sum - сумма в колонке,
        - Average - среднее арифметическое колонки,
        - Average X - конвертация в GRADE_CONVERTED систему.
        Параметр конвертации для AverageInFive можно изменить по переменной 
        GRADE_CONVERTED = 5;

        Args:
            - nameHeadersColumns_Sum_Average_Round - список названия колонок. 
            При желании названия по умолчанию можно изменить.
            
            - nameHeadersString_Max_Sum_Average - список названия строк. При 
            желании названия по умолчанию можно изменить.

        Return:
            - Таблица pd.DataFrame с новыми колонками и строчками.
        """
        return TableHandler.createTableToViewWith__Sum_Avg_Round(self.__data_table, self.__headersGradesStudents, self.__names, nameColumnStuneds, nameHeadersColumn_Sum_Average_Round, nameHeadersString_Max_Sum_Average)
    
    def createTableGradesTest(self, 
                              nameHeadersColumns_Sum_Average_Round:List[str] = ["Sum", f"Average {GRADE_CONVERTED}", f"Round {ROUND_FACTOR}"]
                              ) -> pd.DataFrame:
        """
        Создаёт таблицу оценок тестов pd.DataFrame, добавляя к ней колонки Sum,
        Average, Round.
        - Sum - сумма всех чисел строки
        - Average X - округление числа в X оценку, находя при этом среднее
        - Round X - округление числа с заданным ROUND_FACTOR. Параметр 
        округляет число с заданной точностью.

        Args:
            - nameHeadersColumns_Sum_Average_Round - список названия колонок.
            При желании названия по умолчанию можно изменить.

        Return:
            - Таблица pd.DataFrame с новыми колонками.
        """
        return TableHandler.createTableWithNewColumns_SumAverageRound(self.__data_table, 
                                                                      self.__headersTestScore, 
                                                                      nameHeadersColumns_Sum_Average_Round)

    
    def createTableGradesTestToView(self, 
                                    nameColumnStuneds: str = "Students",
                                    nameHeadersColumn_Sum_Average_Round:List[str] = ["Sum", f"Average {GRADE_CONVERTED}", f"Round {ROUND_FACTOR}"],  
                                    nameHeadersString_Max_Sum_Average:List[str] = ['Max', 'Sum', 'Average', f"Average {GRADE_CONVERTED}"]
                                    ) -> pd.DataFrame:
        """
        Создаёт таблицу оценок студентов pd.DataFrame, добавляя к ней колонки 
        Sum, Average, Round.
        - Sum - сумма всех чисел строки
        - Average X - округление числа в GRADE_CONVERTED оценку, находя при 
        этом среднее
        - Round X - округление числа с заданным ROUND_FACTOR. Параметр 
        округляет число с заданной точностью. Также добалвяются строки таблицы
        'Max' 'Sum' 'Average' 'AverageInFive':
        
        - Max - максимальное число в колонке,
        - Sum - сумма в колонке,
        - Average - среднее арифметическое колонки,
        - Average X - конвертация в GRADE_CONVERTED систему.
        
        Параметр конвертации для AverageInFive можно изменить по переменной GRADE_CONVERTED = 5;

        Args:
            - nameHeadersColumns_Sum_Average_Round - список названия колонок.
            При желании названия по умолчанию можно изменить.
            
            - nameHeadersString_Max_Sum_Average - список названия строк. При 
            желании названия по умолчанию можно изменить.

        Return:
            - Таблица pd.DataFrame с новыми колонками и строчками.
        """
        return TableHandler.createTableToViewWith__Sum_Avg_Round(self.__data_table, self.__headersTestScore, self.__names, nameColumnStuneds, nameHeadersColumn_Sum_Average_Round, nameHeadersString_Max_Sum_Average)
    
    
    def export_PngPieBRSO(self, 
                          fileToExport: str = "BRSOpie.png",
                          nameHeader = 'Средняя успеваемость по БРСО',
                          colors: List[str] =['c', 'moccasin', 'sienna', 
                                              'silver', 'gold']
                          ) -> None:
        """
        По умолчанию экспортирует пирожковую диаграмму BRSO в эту папку с 
        названием BRSOpie.png, где находится программа. При необходимости 
        параметр fileToExport можно изменить.
        
        Raise:
            - Файл {fileToExport} уже существует      

        Args:
            - fileToExport - название файла и место куда его нужно экспортировать
            (.png). По умолчанию BRSOpie.png
            
            - nameHeader - название заголовка диаграммы. По умолчанию 
            "Средняя успеваемость по БРСО"
            - colors: список названия цветов. Берутся из библиотеки plt. По 
            умолчанию уже стоят.

        Return:
            - None;
        """
        if os.path.isfile(fileToExport):
                raise FileExistsError(f"Файл {fileToExport} уже существует")
            
        roundGrades = self.createTableGradesStudents().iloc[:, -1].to_list()

        gradesSet = set(roundGrades)
        
        labels = {}
        
        for grade in gradesSet:
            labels[grade] = roundGrades.count(grade)

        fig, ax = plt.subplots()
        
        ax.pie(labels.values(), labels=labels.keys(), colors=colors, autopct='%1.1f%%', startangle=140)

        ax.set_title(nameHeader)

        ax.legend(loc='upper right')

        ax.axis('equal')
        fig.savefig(fileToExport)

        return None
    
    def export_PngPieOTS(self, 
                         fileToExport: str = "OTSpie.png", 
                         nameHeader = 'Оценка тестов ОТС', 
                         colors: List[str] =['c', 'moccasin', 'sienna', 'silver', 'gold']
                         ) -> None:
        """
        По умолчанию экспортирует пирожковую диаграмму OTS в эту папку с 
        названием OTSpie.png, где находится программа. При необходимости 
        параметр fileToExport можно изменить.
        
        Raise:
            - "Файл {fileToExport} уже существует"
        Args:
            - fileToExport - название файла и место куда его нужно экспортировать
            (.png). По умолчанию OTSpie.png
            - nameHeader - название заголовка диаграммы. По умолчанию 
            "Оценка тестов ОТС"
            - colors: список названия цветов. Берутся из библиотеки plt. По 
            умолчанию уже стоят
            
        Return:
            - None;
        """
        
        if os.path.isfile(fileToExport):
            raise FileExistsError(f"Файл {fileToExport} уже существует")
        
        roundGrades = self.createTableGradesTest().iloc[:, -1].to_list()

        gradesSet = set(roundGrades)
        
        labels = {}
        
        for grade in gradesSet:
            labels[grade] = roundGrades.count(grade)

        fig, ax = plt.subplots()
        
        ax.pie(labels.values(), labels=labels.keys(), colors=colors, autopct='%1.1f%%', startangle=140)

        ax.set_title(nameHeader)

        ax.legend(loc='upper right')

        ax.axis('equal')
        
        fig.savefig(fileToExport)
        
        return None

    def createTableLtiLsti(self, 
                           nameColumnStudent: str = "Students", 
                           nameHeaders_LSI_LTI: List[str] = ["LSI", "LTI"]):
        """
        Создаст таблицу LSI и LTI.
        
        Args:
            - nameColumnStudent: str = "Students" - имя для колонки студентов
            - nameHeaders_LSI_LTI: List[str] = ["LSI", "LTI"] имена заголовков
            
        Raise:
            - 1. Количество заголовков не 2
        Return:
            - pd.DataFrame - таблица pandas;
        """
        lessimetria = nameHeaders_LSI_LTI

        if(len(lessimetria) != 2):
            raise BadNameHeaders("Количество заголовков не верно, нужно 2, у тебя" + str(len(lessimetria)))

        table_1 = self.__data_table[self.__headersGradesStudents]
        table_2 = self.__data_table[self.__headersTestScore]

        table_1.fillna(0, inplace=True)
        table_1 = table_1.apply(lambda x: pd.to_numeric(x, errors='coerce')).fillna(0)
        table_1 = table_1.replace([np.inf, -np.inf], 0)

        table_2.fillna(0, inplace=True)
        table_2 = table_2.apply(lambda x: pd.to_numeric(x, errors='coerce')).fillna(0)
        table_2 = table_2.replace([np.inf, -np.inf], 0)



        headersTableLtiLsi = [str(x) for x in range(1, len(self.__headersGradesStudents) + 1)]

        questionsTable1Dict = dict(zip(headersTableLtiLsi, self.__headersGradesStudents))

        questionsTable2Dict = dict(zip( headersTableLtiLsi, self.__headersTestScore))


        table_3 = pd.DataFrame()
        
        for name in headersTableLtiLsi:
            table_3[name] = table_2[questionsTable2Dict[name]] / table_1[questionsTable1Dict[name]]



        table_3.fillna(0, inplace=True)
        table_3 = table_3.apply(lambda x: pd.to_numeric(x, errors='coerce')).fillna(0)
        table_3 = table_3.replace([np.inf, -np.inf], 0)


        table_3[lessimetria[0]] = table_3.apply(TableHandler.calculateAverage, axis=1)
        table_3 = table_3.transpose()
        table_3[lessimetria[1]] = table_3.apply(TableHandler.calculateAverage, axis=1)
        table_3 = table_3.transpose()

        table_3[nameColumnStudent] =  self.__names + [lessimetria[1]]
        table_3 = table_3[[table_3.columns[-1]] + list(table_3.columns[:-1])].reset_index(drop=True)        
        
        return table_3

    def export_PngPopularityTests(self, 
                                  fileToExport: str = "popularityTests.png",
                                  namesHeader: List[str] = ['Оценка популярности заданий', 'Виды заданий', 'Средние баллы'],
                                  colorBars: str = 'green') -> None:
        """
        По умолчанию экспортирует диаграмму популярности тестов файл с названием
        popularityTests.png, где находится программа. При необходимости параметр 
        fileToExport можно изменить.

        Raise:
            - Файл {fileToExport} уже существует.
            - Количество заголовков не верно для графика, нужно 3.
        
        Args:
            - fileToExport - название файла и место куда его нужно экспортировать
            (.png). По умолчанию popularityTests.png
            - namesHeader - название заголовков диаграммы. По умолчанию 'Оценка
            популярности заданий', 'Виды заданий', 'Средние баллы'
            - colorBars: - цвета колонок в диаграмме. По умолчанию 'green'. 
            Берутся из plt.
            
        Return:
            - None;
        """
        if os.path.isfile(fileToExport):
            raise FileExistsError(f"Файл {fileToExport} уже существует")
        
        if(len(namesHeader) != 3):
            raise BadNameHeaders("Количество заголовков не верно для графика, нужно 3, у тебя" + str(len(namesHeader)))
        
        tmpTable = self.createTableGradesTestToView()
        tmpTable = tmpTable[self.__headersTestScore].transpose()
        
        x = tmpTable.iloc[:, -1]
        x = x.to_list()
        
        y = [i for i in range(1, len(x) + 1)]
        
        fig, ax = plt.subplots()
        
        ax.bar(y, x, width=0.5, color=colorBars)

        ax.set_xlabel(namesHeader[1])
        ax.set_ylabel(namesHeader[2])
        ax.set_title(namesHeader[0])

        ax.grid(True, axis='y')

        fig.savefig(fileToExport)
        
        return None
    
    def export_PngMotivation(self, 
                             fileToExportEducationMotivation: str = "edu_mot.png",
                             fileToExportMotivationEducation: str = "mot_edu.png",
                             namesHeaders: List[str] = ['Соотношение успеваемости к мотивации в группе', 'Успеваемость', 'Мотивация'],
                             colorPoint: str = "salmon",
                             colorLine: str = "k") -> None:
        """
        По умолчанию экспортирует диаграммы мотивации и успеваемости в файл с 
        названием popularityTests.png, где находится программа. При необходимости 
        параметры места экспорта можно изменить.
                
        Args:
            - fileToExportEducationMotivation: str = "edu_mot.png" - название файла
            и место куда экспортировать первый график.
            - fileToExportMotivationEducation: str = "mot_edu.png" - название файла
            и место куда экспортировать второй график.
            - namesHeaders: List[str] - имена заголовков для графиков.
            - colorPoint: str = "salmon" - цвет точек на графике.
            - colorLine: str = "k" - цвет линии графика.
            
        Return:
            - None;
        """
        if os.path.isfile(fileToExportEducationMotivation) or os.path.isfile(fileToExportMotivationEducation):
            raise FileExistsError(f"Файл {fileToExportEducationMotivation} или {fileToExportMotivationEducation} уже существует.")
        if(len(namesHeaders) != 3):
            raise BadNameHeaders("Количество заголовков не верно, нужно 3, у тебя" + str(len(namesHeaders)))
        
        data_x = self.createTableGradesStudents().iloc[:, -2].to_list()
        data_y = self.createTableGradesTest().iloc[:, -2].to_list()


        x = np.array(data_x)
        y = np.array(data_y)

        A = np.vstack([x, np.ones(len(x))]).T
        m, c = np.linalg.lstsq(A, y, rcond=None)[0]

        fig, ax = plt.subplots()
        
        ax.scatter(x, y, color=colorPoint, label='Точки', marker='D')
        ax.plot(x, m*x + c, color=colorLine, label='Прямая')

        ax.set_title(namesHeaders[0])
        ax.set_xlabel(namesHeaders[2])
        ax.set_ylabel(namesHeaders[1])
        ax.legend()
        ax.grid(True, axis='y')

        fig.savefig(fileToExportEducationMotivation)

        fig, ax = plt.subplots()

        x = np.array(data_y)
        y = np.array(data_x)

        A = np.vstack([x, np.ones(len(x))]).T
        m, c = np.linalg.lstsq(A, y, rcond=None)[0]

        ax.scatter(x, y, color=colorPoint, label='Точки', marker='D')
        ax.plot(x, m*x + c, color=colorLine, label='Прямая')

        ax.set_title(namesHeaders[0])
        ax.set_xlabel(namesHeaders[1])
        ax.set_ylabel(namesHeaders[2])
        
        ax.legend()
        ax.grid(True, axis='y')

        fig.savefig(fileToExportMotivationEducation)
        
        return None 
    
    
    CONST_STEP = 4
    
    def export_PngBenefits(self, 
                           fileToExport: str = "benefits.png",
                           headerBenefitsQuestion: str = "17. Какими средствами обучения вы преимущественно пользовались?",
                           titleAndLabels: List[str] = ["Пособия", "Количество", "Пособие"],
                           typesBenefitsForPng: List[str] = ['Электронные учебники', 'Рабочие тетради', 'Видеолекции', 'Печатные учебники'],
                           typesBenefitsInTable: List[str] = ['Электронными учебниками', 'Рабочими тетрадями', 'Видеолекциями', 'Печатными учебниками']) -> None:
        """
        По умолчанию экспортирует диаграмму пособия файл с названием benefits.png, 
        где находится программа. При необходимости параметр место экспорта можно 
        изменить.
        
        Raise:
            - Файл {fileToExport} уже существует
            - Количество уникальныъ значений typesBenefitsForPng и typesBenefitsInTable должно быть одинаковым
            - Количество названий в titleAndLabels неправильное. Должно быть 3
        CONST_STEP = 4, параметр для шага в графике по оси y
        Args:
            - fileToExport: str = "benefits.png" - название файла и место куда 
            экспортировать график.
            - headerBenefitsQuestion: str = "17. Какими средствами обучения вы 
            преимущественно пользовались?" - название заголовка с пособиями
            - titleAndLabels: List[str] = ["Пособия", "Количество", "Пособие"], - параметры для отрисовки графика.
            - typesBenefitsForPng: List[str] = ['Электронные учебники', 'Рабочие
            тетради', 'Видеолекции', 'Печатные учебники'] - названия колонок для
            графика
            - typesBenefitsInTable: List[str] = ['Электронными учебниками', 
            'Рабочими тетрадями', 'Видеолекциями', 'Печатными учебниками'] - названия пособий в таблице данных. Нужно соотнести с параметром выше.
        Return:
            - None;
        """
        
        if os.path.isfile(fileToExport):
            raise FileExistsError(f"Файл {fileToExport} уже существует")
        
        if len(set(typesBenefitsInTable)) != len(set(typesBenefitsForPng)) or len(typesBenefitsForPng) != len(set(typesBenefitsInTable)):
            raise BadNameHeaders("typesBenefitsForPng и typesBenefitsInTable должны быть одинаковыми")
        
        if len(titleAndLabels) != 3:
            raise BadNameHeaders("Количество названий в titleAndLabels неправильное. Должно быть 3, а у тебя" + len(titleAndLabels))
        
        dfBenefits = self.__data_table[headerBenefitsQuestion]


        fig, ax = plt.subplots()
        
        benefitsDict = dict(zip(typesBenefitsInTable, typesBenefitsForPng))
        benefitsDict2 = dict()
        
        
        for i in typesBenefitsForPng:
            benefitsDict2[i] = 0
            

        for i in dfBenefits:
            for word in typesBenefitsInTable:
                if word in i:
                    benefitsDict2[benefitsDict[word]] += 1


        x = list(benefitsDict2.keys())
        y = list(benefitsDict2.values())

        fig.set_size_inches(10, 5)
        ax.bar(x, y, width=0.5)
        
        ax.set_yticks(range( max(y) + TableHandler.CONST_STEP)[::TableHandler.CONST_STEP])
        
        ax.set_xlabel(titleAndLabels[2])
        ax.set_ylabel(titleAndLabels[1])
        ax.set_title(titleAndLabels[0])

        ax.grid(True, axis='y')

        fig.savefig(fileToExport)
        
        return None
        

    @staticmethod
    def createTableToViewWith__Sum_Avg_Round(tableValues: pd.DataFrame, 
                                             headersForCalculation: List[str],
                                             namesStudents: List[str],
                                             nameColumnStuneds: str = "Students",
                                             nameHeadersColumn_Sum_Average_Round:List[str] = ["Sum",  f"Average {GRADE_CONVERTED}", f"Round {ROUND_FACTOR}"],  
                                             nameHeadersString_Max_Sum_Average:List[str] = ['Max', 'Sum', 'Average', f"Average {GRADE_CONVERTED}"]) -> pd.DataFrame:
        """
        Создаёт таблицу pd.DataFrame, добавляя к ней колонки Sum, Average, Round.
        - Sum - сумма всех чисел строки
        - Average X - округление числа в GRADE_CONVERTED оценку, находя при 
        этом среднее
        - Round X - округление числа с заданным ROUND_FACTOR = 0.5. Параметр
        округляет число с заданной точностью.
        Также добалвяются строки таблицы 'Max' 'Sum' 'Average' 'AverageInFive' 
        - Max - максимальное число в колонке,
        - Sum - сумма в колонке,
        - Average - среднее арифметическое колонки,
        - Average X - конвертация в GRADE_CONVERTED систему.
        Параметр конвертации для AverageInFive можно изменить по переменной
        GRADE_CONVERTED = 5;

        Args:
            - tableValues: pd.DataFrame - таблица для обработки.
            - headersForCalculation: List[str] - список вопросов для вычислений.
            - namesStudents: List[str] - Колонка студенты.
            - nameColumnStuneds: str = "Students" - Имя колонки 'Студенты'
            - nameHeadersColumns_Sum_Average_Round:List[str] = ["Sum", 
            "Average Grade", "Round Grade"] - названия для Sum, Average Grade, Round Grade.
            - nameHeadersString_Max_Sum_Average:List[str] = ['Max', 'Sum', 
            'Average', 'AverageInFive']) - названия для 'Max', 'Sum', 'Average', 
            'AverageInFive'.
        Return:
            - Таблица pd.DataFrame с новыми колонками и строчками.
        """
        if len(nameHeadersString_Max_Sum_Average) != 4:
            raise BadNameHeaders("Количество имён заголовков для суммы, среднего и округления должно быть 4, а у тебя" + len(nameHeadersString_Max_Sum_Average))
        
        gradeStudents = TableHandler.createTableWithNewColumns_SumAverageRound(tableValues, headersForCalculation, nameHeadersColumn_Sum_Average_Round)

        gradeStudentsResult_Headers = headersForCalculation + nameHeadersColumn_Sum_Average_Round

        gradeStudents_MaxSumAverage = pd.DataFrame(columns=nameHeadersString_Max_Sum_Average)

        for column in gradeStudentsResult_Headers:
            col_values = gradeStudents[column]

            max_val = max(col_values)
            sum_val = sum(col_values)
            avg_val = sum_val / len(col_values)
            avg_five = TableHandler.GRADE_CONVERTED * avg_val / max_val
            gradeStudents_MaxSumAverage.loc[column] = [max_val, sum_val, avg_val, avg_five]

        gradeStudents_MaxSumAverage = gradeStudents_MaxSumAverage.transpose()

        gradeStudentsResult = pd.concat([gradeStudents, gradeStudents_MaxSumAverage])
        
        gradeStudentsResult[nameColumnStuneds] = namesStudents + nameHeadersString_Max_Sum_Average
        
        gradeStudentsResult = gradeStudentsResult[[gradeStudentsResult.columns[-1]] + list(gradeStudentsResult.columns[:-1])].reset_index(drop=True)
        
        return gradeStudentsResult
    
    @staticmethod
    def createTableWithNewColumns_SumAverageRound(tableValues: pd.DataFrame, 
                                                  headersForCalculation: List[str],
                                                  nameHeadersColumns_Sum_Average_Round:List[str] = ["Sum",  f"Average {GRADE_CONVERTED}", f"Round {ROUND_FACTOR}"]) -> pd.DataFrame:
        """
        Создаёт таблицу pd.DataFrame, добавляя к ней колонки Sum, Average, Round.
        - Sum - сумма всех чисел строки
        - Average X - округление числа в GRADE_CONVERTED оценку, находя при этом
        среднее
        - Round - округление числа с заданным ROUND_FACTOR = 0.5. Параметр
        округляет число с заданной точностью.

        Args:
            - tableValues: pd.DataFrame - таблица для обработки.
            - headersForCalculation: List[str] - список вопросов для вычислений.
            - nameHeadersColumns_Sum_Average_Round:List[str] = ["Sum", 
            "Average Grade", "Round Grade"] - названия для Sum, Average Grade,
            Round Grade.

        Return:
            - Таблица pd.DataFrame с новыми колонками.
        """
        if len(nameHeadersColumns_Sum_Average_Round) != 3:
            raise BadNameHeaders("Количество имён заголовков для суммы, среднего и округления должно быть 3, а у тебя" + len(nameHeadersColumns_Sum_Average_Round))

        gradesStudents = tableValues[headersForCalculation]

        gradesStudents.fillna(0, inplace=True)
        gradesStudents = gradesStudents.apply(lambda x: pd.to_numeric(x, errors='coerce')).fillna(0)
        gradesStudents = gradesStudents.replace([np.inf, -np.inf], 0)

        gradesStudents[nameHeadersColumns_Sum_Average_Round[0]] = 0
    
        for question in headersForCalculation:
            gradesStudents[nameHeadersColumns_Sum_Average_Round[0]] += gradesStudents[question]

        max_values = {}
    
        for column in headersForCalculation:
            col_values = gradesStudents[column]
            max_val = max(col_values)
            max_values[column] = max_val

        countQuestions = len(headersForCalculation)
    
        gradesStudents[nameHeadersColumns_Sum_Average_Round[1]] = 0
    
    
        for question in headersForCalculation:
            gradesStudents[nameHeadersColumns_Sum_Average_Round[1]] += gradesStudents[question] / max_values[question] * TableHandler.GRADE_CONVERTED / countQuestions
    
        gradesStudents[nameHeadersColumns_Sum_Average_Round[2]] = gradesStudents[nameHeadersColumns_Sum_Average_Round[1]].apply(TableHandler.customRound)

        return gradesStudents
    
    
    def export_TableGradesStudent(self, 
                                  pathForExport: str = "StudentsGrades.xlsx") -> pd.DataFrame:
        """
        Создаёт таблицу pd.DataFrame. Это вывод из createTableGradesStudentsToView()
        плюс идёт экспорт таблицы в папку pathForExport. Параметры для таблицы все
        берутся по умолчанию из функции

        Args:
            - pathForExport: str = "StudentsGrades.xlsx" - файл для экспорта таблицы.

        Return:
            Таблица pd.DataFrame с колонками и строчками.
        """
        table = self.createTableGradesStudentsToView()
        table.to_excel(pathForExport, index=False)
        return table
        
    def export_TableGradesTest(self, 
                               pathForExport: str = "TestGrades.xlsx") -> pd.DataFrame:
        """
        Создаёт таблицу pd.DataFrame. Это вывод из createTableGradesTestToView() 
        плюс идёт экспорт таблицы в папку pathForExport. Параметры для таблицы все берутся по умолчанию из функции

        Args:
            - pathForExport: str = "TestGrades.xlsx" - файл для экспорта таблицы.

        Return:
            - Таблица pd.DataFrame с колонками и строчками.
        """
        table = self.createTableGradesTestToView()
        table.to_excel(pathForExport, index=False)
        return table
        
    def export_TableLtiLsi(self, 
                           pathForExport: str = "LsiLti.xlsx") -> pd.DataFrame:
        """
        Создаёт таблицу pd.DataFrame. Это вывод из createTableLtiLsti() плюс идёт
        экспорт таблицы в папку pathForExport. Параметры для таблицы все берутся 
        по умолчанию из функции

        Args:
            - pathForExport: str = "LsiLti.xlsx" - файл для экспорта таблицы.

        Return:
            - Таблица pd.DataFrame с колонками и строчками.
        """
        table = self.createTableLtiLsti()
        table.to_excel(pathForExport, index=False)
        return table
    
    def export_TableConclusion(self, 
                               nameFileForExport: str = "Сonclusion.xlsx", 
                               pathForExport: str = "./",
                               nameOriginalForExport: str = "original.xlsx", 
                               exportOriginal: bool = True,
                               nameHeadersColumn_Sum_Average_Round: List[str] = ["Sum", f"Average {GRADE_CONVERTED}", f"Round {ROUND_FACTOR}"],
                               nameHeadersString_Max_Sum_Average: List[str] = ['Max', 'Sum', 'Average', f"Average {GRADE_CONVERTED}"],
                               nameHeaders_LSI_LTI: List[str] = ["LSI", "LTI"]
                               ) -> pd.DataFrame:
        """
        Создаёт таблицу pd.DataFrame. Это вывод из всего анализа научеметрии
        
        nameHeadersColumns_Sum_Average_Round: Создаёт таблицу оценок студентов pd.DataFrame, добавляя к ней колонки Sum, Average, Round.
        - Sum - сумма всех чисел строки
        - Average X - округление числа в GRADE_CONVERTED оценку, находя при этом 
        среднее
        - Round X - округление числа с заданным ROUND_FACTOR. Параметр округляет 
        число с заданной точностью.
        nameHeadersString_Max_Sum_Average: Также добалвяются строки таблицы 'Max' 
        'Sum' 'Average' 'Average X' 
        - Max - максимальное число в колонке,
        - Sum - сумма в колонке,
        - Average - среднее арифметическое колонки,
        - Average X - конвертация в GRADE_CONVERTED систему.
        Параметр конвертации для AverageInFive можно изменить по переменной 
        GRADE_CONVERTED = 5;
        
        Args:
            - nameFileForExport: str  = "Сonclusion.xlsx" - файл для экспорта всех
            таблиц
            - pathForExport: str - Папка для экспорта. Если пути не существует, 
            создаст без исключений.
            - nameHeadersColumns_Sum_Average_Round - список названия колонок. При 
            желании названия по умолчанию можно изменить.           
            - nameHeadersString_Max_Sum_Average - список названия строк. При 
            желании названия по умолчанию можно изменить.
            - nameHeaders_LSI_LTI: List[str] = ["LSI", "LTI"] имена LSI, LTI
            

        Return:
            - Таблица pd.DataFrame со всеми выводами
        """
        os.makedirs(pathForExport, exist_ok=True)
        
        nameColumnStudents = self.__headerStudentsName
        
        if exportOriginal:
            fileOriginal = os.path.join(pathForExport, nameOriginalForExport)
            
            if os.path.isfile(fileOriginal):
                raise FileExistsError(f"Файл {fileOriginal} уже существует")
            self.__data_table.to_excel(fileOriginal, index=False)
        
        file = os.path.join(pathForExport, nameFileForExport)
        
        if os.path.isfile(file):
            raise FileExistsError(f"Файл {file} уже существует")
        
        df1 = self.createTableGradesStudentsToView(nameColumnStudents, nameHeadersColumn_Sum_Average_Round, nameHeadersString_Max_Sum_Average)
        df2 = self.createTableGradesTestToView(nameColumnStudents, nameHeadersColumn_Sum_Average_Round, nameHeadersString_Max_Sum_Average)
        df3 = self.createTableLtiLsti(nameColumnStudents, nameHeaders_LSI_LTI)

        empty_df = pd.DataFrame(columns=[' '])

        combined_df = pd.concat([df1, empty_df, df2, empty_df, df3], axis=1)

        combined_df.to_excel(file, index=False)
        
        return combined_df
    
    def export_PngConslission(self, 
                              pathForExport: str = "./", 
                              namesFiles: List[str] = ["BRSO.png", "OTS.png", "Benefits.png", "Popularity.png", "Motivation.png", "Education.png"],
                              headerBenefitsQuestionInDataTable: str = "17. Какими средствами обучения вы преимущественно пользовались?",
                              typesBenefitsForPng: List[str] = ['Электронные учебники', 'Рабочие тетради', 'Видеолекции', 'Печатные учебники'],
                              typesBenefitsInDataTable: List[str] = ['Электронными учебниками', 'Рабочими тетрадями', 'Видеолекциями', 'Печатными учебниками']
                              ) -> None:
        """
        Экспортирует все выводы в указанную папку. Если ничего не указать, 
        экспортирует всё в основную папку.
        
        Args:
            - pathForExport: str - Папка для экспорта
            - namesFiles: List[str] = ["BRSO.png", "OTS.png", "Benefits.png", 
            "Popularity.png", "Motivation.png", "Education.png"] - имена для 
            файлов.
            - headerBenefitsQuestionInDataTable: str = "17. Какими средствами 
            обучения вы преимущественно пользовались?" - название заголовка для 
            пособий в таблице данных(оснавная таблица),
            - typesBenefitsForPng: List[str] = ['Электронные учебники', 'Рабочие 
            тетради', 'Видеолекции', 'Печатные учебники'] - названия пособий для 
            диаграммы,
            - typesBenefitsInDataTable: List[str] = ['Электронными учебниками', 
            'Рабочими тетрадями', 'Видеолекциями', 'Печатными учебниками'] - названия 
            пособий в таблице для поиска, указывать без пробелов и запятых.
        Raise:
            - 1. Названия не уникальны.
            - 2. Какой-нибудь файл уже сущесвует.
        Return:
            - None.
        """
        if(len(set(namesFiles)) != len(namesFiles) or len(namesFiles) != 6):
            raise BadNameHeaders("Названия не уникальны или количество имён не равно шести")
        
        
        os.makedirs(pathForExport, exist_ok=True)
        
        files = []
        
        for i in namesFiles:
            file = os.path.join(pathForExport, i)
            files.append(file)
        
        self.export_PngPieBRSO(files[0])
        
        self.export_PngPieOTS(files[1])
        
        self.export_PngBenefits(files[2], headerBenefitsQuestionInDataTable, typesBenefitsForPng=typesBenefitsForPng, typesBenefitsInTable=typesBenefitsInDataTable)
        
        self.export_PngPopularityTests(files[3])
        
        self.export_PngMotivation(files[4], files[5])
        
        return None
        
    def export_PngConclussionWithoutBenefits(self, 
                              pathForExport: str = "./", 
                              namesFiles: List[str] = ["BRSO.png", "OTS.png", "Popularity.png", "Motivation.png", "Education.png"],
                              ) -> None:
        """
        Экспортирует все выводы в указанную папку. Если ничего не указать, 
        экспортирует всё в основную папку.
        
        Args:
            - pathForExport: str - Папка для экспорта
            - namesFiles: List[str] = ["BRSO.png", "OTS.png", "Benefits.png", 
            "Popularity.png", "Motivation.png", "Education.png"] - имена для 
            файлов.
        Raise:
            - 1. Названия не уникальны.
            - 2. Какой-нибудь файл уже сущесвует.
        Return:
            - None.
        """
        if(len(set(namesFiles)) != len(namesFiles) or len(namesFiles) != 5):
            raise BadNameHeaders("Названия не уникальны или количество имён не равно пяти")
        
        
        os.makedirs(pathForExport, exist_ok=True)
        
        files = []
        
        for i in namesFiles:
            file = os.path.join(pathForExport, i)
            files.append(file)
        
        self.export_PngPieBRSO(files[0])
        
        self.export_PngPieOTS(files[1])
           
        self.export_PngPopularityTests(files[2])
        
        self.export_PngMotivation(files[3], files[4])
        
        return None
        
    @staticmethod
    def getHeadrsExcelToList(pathToFile: str) -> List[str]:
        
        df = pd.read_excel(pathToFile)
        headers = df.columns.tolist()
        
        return headers
         
    
    @property
    def dataTable(self) -> pd.DataFrame:
        return self.__data_table
