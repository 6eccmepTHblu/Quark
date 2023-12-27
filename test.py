from pprint import pprint
from functools import reduce


data = [
    {'Вес общий': '55.3', 'Количество': '1', 'Марка': '11А204-86', 'Наименование': 'Связь горизонтальная', 'ПутьФайла': '...'},
    {'Вес общий': '55.3', 'Количество': '2', 'Марка': '11А204-86', 'Наименование': 'Связь горизонтальная', 'ПутьФайла': '...'},
    {'Вес общий': '55.3', 'Количество': '3', 'Марка': '11А204-82', 'Наименование': 'Связь горизонтальная', 'ПутьФайла': '...'},
    {'Вес общий': '55.3', 'Количество': '4', 'Марка': '11А204-82', 'Наименование': 'Связь горизонтальная', 'ПутьФайла': '...'},
    {'Вес общий': '55.3', 'Количество': '5', 'Марка': '11А204-82', 'Наименование': 'Связь горизонтальная', 'ПутьФайла': '...'},
    {'Вес общий': '55.3', 'Количество': '6', 'Марка': '11А204-86', 'Наименование': 'Связь горизонтальная', 'ПутьФайла': '...'},
    {'Вес общий': '55.3', 'Количество': '7', 'Марка': '11А204-86', 'Наименование': 'Связь горизонтальная', 'ПутьФайла': '...'},
    {'Вес общий': '55.3', 'Количество': '8', 'Марка': '11А204-86', 'Наименование': 'Связь горизонтальная', 'ПутьФайла': '...'},
    {'Вес общий': '55.3', 'Количество': '9', 'Марка': '11А204-88', 'Наименование': 'Связь горизонтальная', 'ПутьФайла': '...'}
]

def sums_rows_by_keys(table: list[dict], keys: list[str], keys_sum: list[str]) -> list[dict]:
    if not table or not keys and not keys_sum:
        return table

    key_name = '__ключ_слияния__'
    for row in table:
        row[key_name] = '!'.join([str(row.get(key, '')) for key in keys])

    set_keys = set(row[key_name] for row in table)

    all_data = []
    for key in set_keys:
        filter_data = [row for row in table if row[key_name] == key]
        if len(filter_data) == 1:
            all_data.extend(filter_data)
        else:
            number = -1
            for row in filter_data:
                result_sum = True
                for key_sum in keys_sum:
                    try:
                        result_sum *= isinstance(float(str(row.get(key_sum, '+'))), float)
                    except:
                        result_sum *= False
                        break
                if result_sum:
                    if number == -1:
                        all_data.extend([row])
                        number = len(all_data)-1
                    else:
                        for key_sum in keys_sum:
                            all_data[number][key_sum] = str(float(all_data[number][key_sum]) + float(row[key_sum]))
                else:
                    all_data.extend([row])
    return all_data


table = sums_rows_by_keys(data, ['Марка'], ['Количество', 'Вес общий'])
