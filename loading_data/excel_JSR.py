import openpyxl
import fnmatch
import re

from def_folder import data_collection as fun
from def_folder import data_normalization as norm
from def_folder import excel_collection as excel

TEMPLATES = ['*ЖСР*.xls*',
             '*Журнал сварочных работ*.xls*']
HEADERS = {'Дата работ': 'Дата выполнения работ*',
           'Свар.Элемент': 'Место или номер*',
           'Наим/Сталь': 'Наименование соединяемых элементов*',
           'Клеймо': 'Клеймо',
           'Заключения': 'Отметка о приемке*'}
SHEET_REESTR = '*основа*'
RESULT = ['не годен',
          'годен',
          'вырез',
          'ремонт']


def get_data(path):
    # Найти файл
    file = fun.find_file(path, TEMPLATES)
    if file is None:
        return None

    # Открываем книгу АООК
    wb = openpyxl.load_workbook(file)

    # Собираем реестры АООК
    reestr_jsr = []
    for sh in wb:
        if fnmatch.fnmatch(sh.title.lower(), SHEET_REESTR):
            all_data = excel.get_data_excel(sh)

            # Ищем нужные столбцы
            headers, resul_rows = excel.find_headers(all_data, HEADERS)

            # Выбераем нужные столбцы, разбивая по словарю
            data = []
            for row in all_data[resul_rows:]:
                row_data = {}
                for key, item in headers.items():
                    row_data[key] = row[item]
                data.append(row_data)

            # Фильтруем данные по АоРПИ
            for row in data:
                if re.search(r"(\d+\.?\d+\.?\d+)", str(row['Дата работ'])):
                    reestr_jsr.append(row)

    # Разделяем данные по заключениям
    split_jsr = norm.splits_by_character(reestr_jsr, 'Заключения', '.,')

    # Нормирование
    for row in split_jsr:
        if 'Клеймо' in row:
            row['Клеймо'] = norm.transliteration_ru_en(row['Клеймо'])
        if 'Дата работ' in row:
            row['Дата работ'] = norm.get_date(row['Дата работ'])
        if 'Наим/Сталь' in row:
            row['Наим/Сталь'] = norm.transliteration_ru_en(row['Наим/Сталь']).upper()
        if 'Свар.Элемент' in row:
            row['Свар.Элемент'] = row['Свар.Элемент'].replace('№', '')
        if 'Заключения' in row:
            # Тип
            if fnmatch.fnmatch(row['Заключения'], '*ВИК *'):
                row['Тип'] = 'ВИК Лимак'
            elif fnmatch.fnmatch(row['Заключения'], '*каппилярного*'):
                row['Тип'] = 'ПВК Лимак'

            # Дата
            date = row['Заключения'].split(' от ')
            if len(date) > 1:
                date = norm.get_date(date[1])
            else:
                date = norm.get_date(row['Заключения'])
            if not date:
                date = ''
            row['Дата заключения'] = date

            # Результат
            row['Результат заключения'] = ''
            for result in RESULT:
                if fnmatch.fnmatch(row['Заключения'].lower(), '*' + result + '*'):
                    row['Результат заключения'] = result
                    break

            # Номер
            row['Номер заключения'] = ''
            nambe = row['Заключения'].split(' от ')[0].split('№')
            if len(nambe) > 1:
                nambe = nambe[1]
            else:
                nambe = nambe[0]
            row['Номер заключения'] = nambe.replace(' ', '')

    print('Данные из ЖСР собранны - ' + str(len(split_jsr)) + '.')
    return split_jsr


if __name__ == '__main__':
    get_data('D:\Работа\Силенко Д.Т\Задача 2(Путхон - Кварк V2)\Данные')
