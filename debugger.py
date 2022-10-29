# LTC_1_notebook/notebooks/FinderData.ipynb
# importing some libraries
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


row_razdel, start_row = searcher_row_razdel(data_excel)


def make_rows(data, start_row):

    Row_list = []

    for index, rows in data[start_row + 1:].iterrows():
        my_list = [rows.Col0, rows.Col1, rows.Col2, rows.Col3,
                   rows.Col4, rows.Col5, rows.Col6, rows.Col7,
                   rows.Col8, rows.Col9, rows.Col10, rows.Col11,
                   rows.Col12, rows.Col13, rows.Col14, rows.Col15,
                   rows.Col16, rows.Col17, rows.Col8, rows.Col19,
                   rows.Col20]

        Row_list.append(my_list)

    return Row_list


rows_ = make_rows(data=data_excel, start_row=start_row)


def deleter(rows):
    for i in rows[:17]:
        # print(i)
        for j in i:
            if str(j) == 'nan':
                del i[j]
            else:
                continue
    return rows


full = deleter(rows=rows_[:16])
