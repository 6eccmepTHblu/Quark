import os
import traceback
import logging
import re

from datetime import date, datetime
from sys import argv

from def_folder import data_collection as fun
from def_folder import normalization as norm
from def_folder import data_normalization
from data_output import out__report as out
from def_folder.data_collection import types_date
from tkinter import simpledialog

try:
    script, path = argv
except:
    path = r'D:\Работа\Силенко Д.Т\Задача 2(Путхон - Кварк V2)\Авансовый'

FAIL_NAME = os.path.join(path, "Отчет.xlsx")

# Сверка АОРПИ(+АООК) с КМД___________________________________________________
def reconciles_documents_report():
    # Получаем данные заключений из АОРПИ
    data_report = fun.collects_data_by_type(path, ['Авансовый отчет'], norml=False)
    print('Данные из отчета - ' + str(len(data_report)) + '.')

    # Создание основного окна
    number = simpledialog.askinteger("Input", "Введите число:")

    # Оставляем только данные с итогом
    for i, row in enumerate(data_report):
        row['key'] = i
    data_report_result = data_report[:]
    data_report_result = norm.normalisation_par_type_de_fichier(data_report_result, 'Авансовый отчет таблица')

    # Распределение данных по итогам
    for i, row in enumerate(data_report_result):
        list_data = []
        row['ТаблицаИтого'] = ''
        for row_2 in data_report:
            if row['ПутьФайла'] == row_2['ПутьФайла']:
                list_data.append(row_2)

        for row_2 in list_data:
            if row['key'] != row_2['key'] or len(list_data) == 1:
                if row_2['Сумма'] != '5':
                    row['ТаблицаИтого'] = data_normalization.append_value(
                        row['ТаблицаИтого'],
                        f'@@@{row_2["Номер"]}@@{row_2["Дата"]}@@{row_2["Описание"]}@@{row_2["Сумма"]}@@@')

    # Дополнить данные для вывода
    date_col_name = ['Дата 1']
    for i, row in enumerate(data_report_result):
        row.setdefault('Дата поступления', date.today().strftime('%d.%m.%Y'))
        date_list = []
        for date_input in date_col_name:
            date_re = re.search(r'\d\d\.\d\d\.\d\d\d\d', row.get(date_input))
            if date_re:
                date_re = types_date(date_re.group(0))
                if isinstance(date_re, datetime):
                    date_list.append(date_re)
        if len(date_list) == 0:
            row['Дата прибытия'] = row.get('Дата 1', '')
        elif len(date_list) == 1:
            row['Дата прибытия'] = date_list[0].strftime("%d.%m.%Y")
        else:
            row['Дата прибытия'] = max(date_list).strftime("%d.%m.%Y")
        mouth = row['Дата прибытия'].split('.')
        row.setdefault('Номер удостоверения', str(mouth[1] if len(mouth) > 1 else 0) + '/' + str(number+i))

    # Вывод данныех в Excel
    out.data_output(data_report_result, FAIL_NAME)


def main():
    try:
        # Удаляем файл, если он етсь
        if os.path.exists(FAIL_NAME):
            os.remove(FAIL_NAME)

        # # Выполняем сверки
        reconciles_documents_report()
    except Exception as ex:
        # Настройка конфигурации логирования
        logging.basicConfig(filename=path + '\\Отчет об ошибках.txt', level=logging.ERROR,
                            format='%(asctime)s - %(levelname)s - %(message)s')
        print('Ошибка:\n', traceback.format_exc())
        logging.exception("Произошла ошибка: %s", str(ex))
        return exit(1)


if __name__ == "__main__":
    main()
