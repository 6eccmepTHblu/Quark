import logging

import pandas as pd

from def_folder.data_normalization import create_hyperlink

STATUS = {'АОРПИ !=': 'АоРПИ не указан в АООК/АОСР',
          'АООК !=': 'Не приложен АОРПИ',
          'АООК != дата': 'Дата АоРПИ не равна дате АоРПИ в АООК/АОСР',
          'Совпало': 'Полное совпадение',
          'АООК < АоРПИ': 'Дата АООК раньше АоРПИ'}  # Итоговые статусы проверки


def reconciliation_data(aorpi, aook, date_aook):
    if not aorpi:
        logging.warning(f"Сверка документов АООК с АоРПИ не проведена! Данные по АоРПИ отсутствуют!")
        return None

    if not aook:
        logging.warning(f"Сверка документов АООК с АоРПИ не проведена! Данные по АООК отсутствуют!")
        return None

    # Объединение по ключу
    data_itog = pd.merge(
        pd.DataFrame(aorpi),
        pd.DataFrame(aook),
        left_on='Номер',
        right_on='АоРПИ Номер из АООК',
        how='outer',
        suffixes=('_1', '_2'),
        validate="many_to_many"
    )

    # Преобразуем даты в даты
    data_itog['Дата АООК'] = pd.to_datetime(date_aook, dayfirst=True)
    data_itog['_Дата'] = pd.to_datetime(data_itog['Дата'], format='%d.%m.%Y', errors='coerce')
    data_itog['_Дата АООК'] = pd.to_datetime(data_itog['Дата АООК'], format='%d.%m.%Y', errors='coerce')
    data_itog['_АоРПИ Дата из АООК'] = pd.to_datetime(data_itog['АоРПИ Дата из АООК'], format='%d.%m.%Y',
                                                      errors='coerce')

    # Конвертируем данные в лист словарей
    data_itog['Расхождения'] = data_itog.apply(lambda row: [], axis=1)
    data_itog = data_itog.map(lambda x: '' if pd.isna(x) else x)
    data_itog = data_itog.to_dict(orient='records')

    # Сверяем даты между собой
    for row in data_itog:
        row['Акт проверки'] = []

        if row['Номер'] == '':
            row['Акт проверки'].append(f"Не приложен АоРПИ №{row['АоРПИ Номер из АООК']} из АООК.")
            row['Статус проверки'] = STATUS['АООК !=']
        elif row['АоРПИ Номер из АООК'] == '':
            row['Акт проверки'].append(f"АоРПИ №{row['Номер']} не указан в АООК.")
            row['Статус проверки'] = STATUS['АОРПИ !=']
        elif row['_Дата'] != row['_АоРПИ Дата из АООК']:
            row['Акт проверки'].append(
                f"Дата {row['АоРПИ Дата из АООК']} АОРПИ №{row['АоРПИ Номер из АООК']} не равна дате {row['Дата']} АОРПИ в АООК.")
            row['Статус проверки'] = STATUS['АООК != дата']
            row['Расхождения'].append({'Тип': 'Дата',
                                       'Рез': row['АоРПИ Дата из АООК'],
                                       'Преф': 'Дата из АООК/АОСР: \n'})
        elif row['_Дата'] > row['_Дата АООК']:
            row['Акт проверки'].append(f"ДК №{row['Номер документа']} не указан в АоРПИ.")
            row['Статус проверки'] = STATUS['АООК < АоРПИ']
            row['Расхождения'].append({'Тип': 'Дата',
                                       'Рез': f"{row['Дата']} > {row['Дата АООК']}",
                                       'Преф': ''})
        else:
            row['Статус проверки'] = STATUS['Совпало']

        if row['Номер'] != '':
            row['Номер'] = create_hyperlink(row["ПутьФайла"], row['Файл'], row['Номер'])
        if row['АоРПИ Номер из АООК'] != '':
            row['АоРПИ Номер из АООК'] = create_hyperlink(row["Путь файла АООК"], '', row['АоРПИ Номер из АООК'])

    logging.info(f"Сверка документов АООК с АоРПИ произведена.")
    return data_itog
