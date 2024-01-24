import os
import xlwings
import getpass
import openpyxl as xl

PATH_FILE = os.path.join(os.getcwd(), 'doc', 'Акт проверки ИД Усть-Луга.xlsm')
MACROS_NAME = 'копирование_листа'


def create_act_checking(path_act: str, number_aook: str, table: list[dict]):
    # Добавляем лист Акта проверки
    vba_book = xlwings.Book(PATH_FILE)
    vba_macro = vba_book.macro(MACROS_NAME)(
        path_act,
        number_aook,
        getpass.getuser(),
        len(table)
    )
    vba_book.close()

    # Заполняем таблицу данными
    wb = xl.load_workbook(path_act)
    sh = wb['Акт проверки']

    for i, elem in enumerate(table, vba_macro):
        sh.cell(row=i, column=2).value = i - vba_macro + 1
        sh.cell(row=i, column=3).value = elem

    wb.save(path_act)

def create_reports(table: dict):
    data_filter = []
    for value in table.values():
        for row in value:
            if row.get('Акт проверки', []):
                data_filter.append(row['Акт проверки'][0])

    return data_filter