# LTC_1_notebook/main.py
# importing libraries

import pandas as pd


class OpenAndSearchTargetColumns:

    """

    Данный класс реализует методы поиска в фале формата .xlsx
    данных, а именно: имена столбцов, которые нам необходимы для шапки
    будущей таблицы и возвращает строку с этими самыми данными


    """
    def __init__(self, path):
        """

        :type path: string
        """
        self.path = path

    def OpenerCouneterRows(self):
        """
        path: str Путь до файла .xlsx
        rerurn: data_exel DataFrame
                ncols list

        Функция возвращает нам DataFrame в виде
        исходного DataFrame с конкретными номерами
        столбцов, т.к. у исходного их нет
        """
        data_excel_head = pd.read_excel(self.path, nrows=50)
        ncols_data = data_excel_head.shape[1]
        ncols = []
        for col in range(ncols_data):
            ncols.append('Col{}'.format(col))

        data = pd.read_excel(self.path, names=ncols)

        return data, ncols

    @staticmethod
    def searcher_row(data):
        """
        data: dataframe
        return: num_row (int)

        Функция ищет строчку, в которой находится №пп для того,
        чтобы потом взять эту строчку и вытянуть из нее необходимую
        шапку для будующей таблицы
        """

        for row, column in data[0:10].iterrows():
            for i in range(len(column)):
                # print(row, column[i])
                if '№пп' in str(column[i]):
                    return row

    @staticmethod
    def target_columns(data, num_row):
        """

        :param data:
        :param num_row:
        :return: string
        Данная функция принимает объект DataFrame и возвращает
        список наименований столбцов в найденной ранее строке.
        """

        string = data.loc[num_row].dropna().to_list()

        return string

    @staticmethod
    def target_columns_list(string):
        """

        :return: target_string str

        Данная функция возвращает только те столбцы, которые нам необходимо
        получить на выходе функции
        """
        target_string = []
        for item in string:
            if item in ['№пп', 'Шифр, номера нормативов и коды ресурсов',
                        'Наименование работ и затрат', 'Ед. изм.', 'Кол-во единиц']:
                target_string.append(item)

        return target_string

    def make_result(self):
        """

        :return: target_string str

        Данная функция подводит итог всего класса, формирующего имена столбцов
        из любого файла разрешения .xlsx, который поступит на вход.

        """
        data_excel, ncols = self.OpenerCouneterRows()
        num_row = self.searcher_row(data=data_excel[:100])
        string = self.target_columns(data=data_excel, num_row=num_row)
        target_string = self.target_columns_list(string=string)
        return target_string


if __name__ == '__main__':
    # some function

    # Путь до файла экселя
    path = 'P:\Python Projects\LTC_1\SN-2012\Глава 1. Здания.xlsx'
    # Экземпляр класса, который решает нашу первую задачу
    OSTC = OpenAndSearchTargetColumns(path=path)
    result = OSTC.make_result()
    print(f'Result of first step is: {result}')


