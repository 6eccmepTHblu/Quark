import logging
import pprint
from fnmatch import fnmatch

from def_folder import data_collection as fun
from def_folder import data_normalization as norm
from def_folder import excel_collection as excel
from def_folder.normalization import normalization_of_data_by_headers

HEADERS = {'Марка': ['*Отправочная марка*','*марка*'],
           'Наименование': ['*Наименование марки*', '*Наименование*'],
           'Количество': ['*Количество*', '*Кол-во*'],
           'Вес общий': ['*Всех*', '*Вес*']}
TEMPLATES = ['*КМД*.xls*']


def get_data(path):
    # Найти файл
    file = fun.find_file(path, TEMPLATES)
    if file == '':
        logging.warning(f"Не найден файл КМД - '{TEMPLATES}'")
        return None

    # Выбераем нужные столбцы, разбивая по словарю
    all_data = excel.get_data_from_sheet(file, HEADERS)

    # Фильтруем данные
    filtered_date = norm.filter_data(all_data, 'Марка', ['', '*Отправочная марка*', '*Количество*'])

    # Нормализация данных по типу
    normal_data = normalization_of_data_by_headers(filtered_date, {'Вес общий': [1,'Вес']}, 'КМД')

    normal_data[:] = [row for row in normal_data if not row['Марка'].lower() == 'марка']
    normal_data[:] = [row for row in normal_data if not row['Марка'].lower() == 'итого']

    # Нормирование
    for row in normal_data:
        if 'Марка' in row:
            row['Марка'] = norm.transliteration_ru_en(row['Марка']).upper().replace(" ", "")
            row['Наименование'] = norm.transliteration_ru_en(row['Наименование'], 'ru').lower()

    logging.info('Данные из КМД собранны - ' + str(len(normal_data)) + '.')
    return normal_data


if __name__ == '__main__':
    pprint.pprint(get_data('D:\Работа\Силенко Д.Т\Задача 2(Путхон - Кварк V2)\Данные'))
