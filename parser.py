import pandas as pd
import re

def OpenerCouneterRows(path):
    """
    path: str Путь до файла .xlsx
    rerurn: data_exel DataFrame
            ncols list

    Функция возвращает нам DataFrame в виде
    исходного DataFrame с конкретными номерами
    столбцов, т.к. у исходного их нет
    """
    data_excel_head = pd.read_excel(path, nrows=50)
    ncols_data = data_excel_head.shape[1]
    ncols = []
    for col in range(ncols_data):
        ncols.append('Col{}'.format(col))

    data_excel = pd.read_excel(path, names=ncols)

    return data_excel, ncols


path = 'P:\Python Projects\LTC_1\SN-2012\Глава 1. Здания.xlsx'

data_excel, ncols = OpenerCouneterRows(path)


def Funck_searcher_fresh_data(data):
    n_row = 1

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
        fresh_data = []

        for row, column in data.iterrows():
            # first
            for i in range(len(column)):
                # print(row, column[i])
                if re.search(r'\bРаздел\b', str(column[i])):
                    razdel_array.append(column[i])
                    # return row
                    n_row = row
            # second
            fresh_data.append(data[n_row])
            return razdel_array, fresh_data, n_row


        return searcher_row_razdel(data)

row_razdel = Funck_searcher_fresh_data(data_excel)

print(row_razdel)