import os
import glob
import pandas as pd
import re


#
class ListOfListsProducer:
    """

    Данный класс реализует функционал открытия файла формата .xlsx,
    конвертирование этого файла в построчное представление,
    выжимку необходимых строк и информации в них и создание
    списка списков в которых будет храниться вся необходимая информация

    """

    def __init__(self, path_to_file):
        """

        :param path_to_file:
        """
        self.path_to_file = path_to_file

    def OpenerCouneterRows(self):
        """
        path: str Путь до файла .xlsx
        rerurn: data_exel DataFrame
                ncols list

        Функция возвращает нам DataFrame в виде
        исходного DataFrame с конкретными номерами
        столбцов, т.к. у исходного их нет
        """
        data_excel_head = pd.read_excel(self.path_to_file, nrows=50)
        ncols_data = data_excel_head.shape[1]
        ncols = []
        for col in range(ncols_data):
            ncols.append('Col{}'.format(col))

        data_excel = pd.read_excel(self.path_to_file, names=ncols)

        return data_excel, ncols

    @staticmethod
    def searcher_row_razdel(data):
        """
            data: dataframe
            return: row (int)

            Функция ищет строчку, в которой находится №пп для того,
            чтобы потом взять эту строчку и вытянуть из нее необходимую
            шапку для будующей таблицы
        """

        global n_row
        razdel_array = []

        for row, column in data.iterrows():
            # first
            for i in range(len(column)):
                # print(row, column[i])
                if re.search(r'\bРаздел\b', str(column[i])):
                    razdel_array.append(column[i])
                    # return row
                    n_row = row

            # second
        return razdel_array, n_row

    @staticmethod
    def make_rows(data, start_row, ncols):
        """

        :param data:
        :param start_row:
        :return:
        """
        Row_list = []

        list_rows = []
        for i in range(len(ncols)):
            list_rows.append(i)

        for index, rows in data[start_row + 1:].iterrows():
            Row_list.append([rows[x] for x in list_rows])

        return Row_list

    def work_with_lists(self):

        """

        Получаем все необходимые данные и используем функции созданные ранее
        return: result (list of lists)
        """
        data_excel, ncols = self.OpenerCouneterRows()

        row_razdel, start_row = self.searcher_row_razdel(data=data_excel)

        rows_ = self.make_rows(data=data_excel, start_row=start_row, ncols=ncols)

        # Создаем списки данных для будущего иерархического списка
        list_of_razdel_nn = []
        list_of_shifrs = []
        list_of_works_janeral = []

        # (rows_[i][4])

        for i in range(len(rows_)):
            # Составляем список подразделов
            if str(rows_[i][0]).isdigit():
                list_of_razdel_nn.append((rows_[i][0]))
            # Составляем список шифров
            if str(rows_[i][2]) != 'nan':
                list_of_shifrs.append(str(rows_[i][2]))
            if str(rows_[i][0]) != 'nan' and str(rows_[i][2]) != 'nan' and str(rows_[i][4]) != 'nan':
                list_of_works_janeral.append(str(rows_[i][4]))

        list_of_works_parts = [[], [], [], []]

        # Дополняем данные
        for i in range(len(rows_)):

            if (str(rows_[i][0]) == 'nan' and
                    str(rows_[i][1]) == 'nan' and
                    str(rows_[i][2]) == 'nan' and
                    str(rows_[i][3]) == 'nan' and
                    str(rows_[i][4]) != 'nan'):

                # Заполняем список частей разработки
                list_of_works_parts[0].append(rows_[i][4])

                # Заполняем спиок единиц измерений частей разработки, если nan, то 'Безразмерная'
                # Для ЗТР если nan то берем клетку по диагонали вверх
                if str(rows_[i][4]) == 'ЗТР':
                    list_of_works_parts[1].append(rows_[i - 1][6])
                elif str(rows_[i][6]) == 'nan' and str(rows_[i][4]) != 'ЗТР':
                    list_of_works_parts[1].append('Безразмерная')
                else:
                    list_of_works_parts[1].append(rows_[i][6])

                # Заполняем список кол-ва единиц, если nan, то 1
                # Для ЗТР если nan то берем клетку по диагонали вверх
                if str(rows_[i][4]) == 'ЗТР':
                    list_of_works_parts[2].append(rows_[i - 1][7])
                elif str(rows_[i][7]) == 'nan':
                    list_of_works_parts[2].append(1)
                else:
                    list_of_works_parts[2].append(rows_[i][7])

                # Заполняем список затрат
                if str(rows_[i][4]) == 'ЗТР':
                    list_of_works_parts[3].append(rows_[i - 1][len(ncols)-2])
                else:
                    list_of_works_parts[3].append(rows_[i][len(ncols) - 5])
                # Для ЗТР если nan то берем клетку по диагонали вверх

            else:
                continue

            if str(rows_[i][4]) == 'ЗТР':
                # Point - метка для разрыва и перехода к другому нормативному документу
                list_of_works_parts[0].append('Point')
                list_of_works_parts[1].append('Point')
                list_of_works_parts[2].append('Point')
                list_of_works_parts[3].append('Point')

        # №пп, Шифр, Наименование работы

        array_for_dict_1 = list(zip(list(list_of_razdel_nn),  # №пп
                                    list(list_of_shifrs),  # Шифр
                                    list(list_of_works_janeral)))  # Наименование работы

        array_for_dict_2 = list(zip(list(list_of_works_parts[0]),  # Наименование работ
                                    list(list_of_works_parts[1]),  # Единицы измерений
                                    list(list_of_works_parts[2]),  # Количество
                                    list(list_of_works_parts[3])))  # Стоимость

        # Создаем упорядоченный список для array_for_dict_2 без точек разрыва
        half_final = []
        lil = []

        # for i in range(len(array_for_dict_2)):
        for i in range(len(array_for_dict_2)):

            if 'Point' not in list(array_for_dict_2[i]):
                lil.append(array_for_dict_2[i])

            elif 'Point' in list(array_for_dict_2[i]):
                # print(lil)
                # print(len(lil))
                half_final.append(lil)
                lil = []
            else:
                continue

        result_list = list(zip(list(array_for_dict_1), list(half_final)))

        return result_list


"""
path = os.getcwd()
exel_files = glob.glob(os.path.join(path, "*.xlsx"))

for f in exel_files:

    # f - путь к файлу
    file = os.open(f, flags=os.O_RDONLY)
    print('Location:', f)
    print('File Name:', f.split("\\")[-1])

    # print the content
    print('Content:')
    print()
"""
print('Start program')
# User input for directory where files to search
expectedDir = '.\\SN-2012'

# Создание директории в которой будут храниться результаты
if not os.path.isdir(".\\results"):
    os.mkdir(".\\results")

"""# Подсчет количества файлов экселя в директории

directory = os.listdir(path=expectedDir)
w = 0
for file in directory:
    if file.endswith((".xlsx", 'xls')):
        w = w + 1
    else:
        continue"""

# first get full file name with directores using for loop
i = 1

for fileName_relative in glob.glob(expectedDir + "**/*.xlsx", recursive=True):
    print("Full file name with directories: ", fileName_relative)
    # Now get the file name with os.path.basename
    fileName_absolute = os.path.basename(fileName_relative)
    print("Only file name: ", fileName_absolute)

    # Экземпляр класса, который решает нашу первую задачу

    LOLP = ListOfListsProducer(path_to_file=fileName_relative)

    try:
        result = LOLP.work_with_lists()
    except AttributeError:
        print('AttributeError проверить функцию подсчета колонок')
        result = 'fileName_relative имеет не то количество колонок'
        continue

    print(f'Получен результат парсинга файла: {fileName_absolute}')
    print(result[1:2])
    save_path = '.\\results'

    print(f'Запись результата парсига {fileName_absolute}')

    completeName = os.path.join(save_path, f'tests_results_{i}' + ".txt")
    with open(completeName, 'w') as f:
        f.write(str(result) + '\n')
        print(f'Запись результата парсига {fileName_absolute} в {completeName} прошла успешно')
    i = i + 1

"""    for i in range(w):

        completeName = os.path.join(save_path, f'tests_results_{i}' + ".txt")
        with open(completeName, 'w') as f:

            f.write(str(result) + '\n')
            print(f'Запись результата парсига {fileName_absolute} в {completeName} прошла успешно')
"""
print('End program')