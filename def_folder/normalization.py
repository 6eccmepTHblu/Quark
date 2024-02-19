import os
import pprint
import re
from fnmatch import fnmatch

from def_folder import data_normalization as norm
from def_folder import data_collection as col
from def_folder.data_normalization import append_value

NOT_SPECIFIED = 'Номер АоРПИ не определен!'

def normalization_of_data_by_headers(table: list[dict], headers: dict, type_name: str|None = None) -> list[dict]:

    if not table or not headers:
        return table

    # Нормирование данных
    for key, item in headers.items():
        if not isinstance(item, list):
            continue

        for check in item[1:]:

            # Загрузка данных для нормирования из csv
            path = os.path.abspath(__file__).replace(os.path.basename(__file__), '')
            path_csv = path + 'normalization\\' + check + '.csv'
            normal_data = []
            if os.path.exists(path_csv):
                normal_data = col.get_data_in_csv([path_csv], name=False, take_strio=False)

            for row in table:
                if key not in row:
                    continue

                # Нормирование данных
                if 'Номер' == check:
                    row[key] = norm.transliteration_ru_en(row[key]).upper()

                elif 'НомерАоРПИ' == check:
                    if row[key] == '':
                        row[key] = row['Номер'] = NOT_SPECIFIED

                elif 'Клеймо' == check:
                    row['Клеймо'] = norm.transliteration_ru_en(row[key])
                    row['Клеймо'] = norm.take_out_the_stamp(row['Клеймо'])

                elif 'Заключение' == check:
                    row[key] = row[key].lower()

                elif 'Материал' == check:
                    row[key] = norm.transliteration_ru_en(row[key]).upper()

                elif 'Марка' == check:
                    row[key] = norm.transliteration_ru_en(row[key]).upper().replace(" ", "")

                elif 'НаимПродукции' == check:
                    row[key] = norm.transliteration_ru_en(row[key], 'ru').lower()

                elif 'Дата' == check:
                    row[key] = row[key].replace(',', '.')
                    row[key] = norm.get_date(row[key])
                    row[key] = norm.remove_side_values(row[key])

                elif 'Количество' == check or 'Сумма' == check:
                    row[key] = re.sub(r'[^0-9.,]', '', row[key])
                    row[key] = norm.remove_side_values(row[key])

                # Замена данных по CSV
                if normal_data:
                    for row_normal in normal_data:
                        to, do = row_normal[0], row_normal[1]
                        if len(row_normal) == 2:
                            type_rep = ''
                        else:
                            type_rep = row_normal[2]
                        if to == '':
                            if row[key] == '':
                                row[key] = do
                        else:
                            if type_rep != '' and row[key] != '':
                                result = ''
                                for text in row[key].split(' '):
                                    if text == to:
                                        result = append_value(result, do, ' ')
                                    else:
                                        result = append_value(result, text, ' ')
                                row[key] = result
                            else:
                                row[key] = row[key].replace(to, do)

                if 'Вес' == check:
                    row[key] = re.sub(r'[^0-9.,]', '', row[key])

                    if fnmatch(row[key], '*.0') and row[key] != '.0':
                        row[key] = row[key][0:len(row[key])-2]

                    try:
                        row[key] = float(row[key])
                        row[key] = round(row[key],2)
                    except ValueError:
                        pass
    return table

def normalisation_par_type_de_fichier(table: list, type_file: str) -> list:

    if not type_file or not table:
        return table

    if 'АОРПИ Усть-Луга сопр.док.' in type_file:
        table[:] = [row for row in table if 'качестве стальных' in row.get('Тип док.', '+')]

    if 'Документ о качестве таблица' in type_file:
        table[:] = [row for row in table if row.get('Марка', '+') != '']
        table[:] = [row for row in table if not fnmatch(row.get('Марка', '').lower(), 'итого*')]
        table[:] = [row for row in table if not fnmatch(row.get('Марка', '').lower(), 'вывод*')]
        table[:] = [row for row in table if not fnmatch(row.get('Марка', '').lower(), 'всего*')]

    if type_file == 'Авансовый отчет таблица':
        table[:] = [row for row in table if row['Сумма'] != '' and row['Сумма'] != '5']

        list_files = set(item.get('ПутьФайла') for item in table)
        result_data = []
        for list_file in list_files:
            list_data = []
            for row in table:
                if list_file == row['ПутьФайла']:
                    list_data.append(row)
            if len(list_data) == 1:
                result_data += list_data
            else:
                filter_data = [row for row in list_data if fnmatch(row['Описание'].lower(), 'ито*') or row['Описание'] == '']
                if len(filter_data) == 1:
                    result_data += filter_data
                else:
                    result_data.append(list_data[-1])
        table[:] = result_data

    return table
