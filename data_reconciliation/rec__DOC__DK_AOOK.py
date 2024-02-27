import logging

import pandas as pd

from def_folder.data_normalization import create_hyperlink

STATUS = {'АОРПИ !=': 'ДК не указан в АООК/АОСР',
          'АООК !=': 'Не приложен ДК',
          'АООК != дата': 'Дата ДК не равна дате ДК в АООК/АОСР',
          'Совпало': 'Полное совпадение',
          'АООК < АоРПИ': 'Дата АООК раньше ДК'}  # Итоговые статусы проверки


def reconciliation_data(dk, aook, date_aook):
    if not dk:
        logging.warning(f"Сверка документов ДК с АООК не проведена! Данные по ДК отсутствуют!")
        return None

    if not aook:
        logging.warning(f"Сверка документов ДК с АООК не проведена! Данные по АООК отсутствуют!")
        return None

    # Объединение по ключу
    data_itog = pd.merge(
        pd.DataFrame(dk),
        pd.DataFrame(aook),
        left_on='Номер документа',
        right_on='Качества Номер из АООК',
        how='outer',
        suffixes=('_1', '_2'),
        validate="many_to_many"
    )
    data_itog = data_itog.drop_duplicates(subset=['Номер документа',
                                                  'Качества Номер из АООК',
                                                  'Дата отгрузки',
                                                  'Качества Дата из АООК'])

    # Преобразуем даты в даты
    data_itog['Дата АООК'] = pd.to_datetime(date_aook, dayfirst=True)
    data_itog['_Дата отгрузки'] = pd.to_datetime(data_itog['Дата отгрузки'], format='%d.%m.%Y', errors='coerce')
    data_itog['_Дата АООК'] = pd.to_datetime(data_itog['Дата АООК'], format='%d.%m.%Y', errors='coerce')
    data_itog['_Качества Дата из АООК'] = pd.to_datetime(data_itog['Качества Дата из АООК'], format='%d.%m.%Y',
                                                         errors='coerce')

    # Конвертируем данные в лист словарей
    data_itog['Расхождения'] = data_itog.apply(lambda row: [], axis=1)
    data_itog = data_itog.map(lambda x: '' if pd.isna(x) else x)
    data_itog = data_itog.to_dict(orient='records')

    # Сверяем даты между собой
    for row in data_itog:
        row['Акт проверки'] = []

        if row['Номер документа'] == '':
            row['Акт проверки'].append(f"Не приложен документ о качестве №{row['Качества Номер из АООК']} из АООК.")
            row['Статус проверки'] = STATUS['АООК !=']
        elif row['Качества Номер из АООК'] == '':
            row['Акт проверки'].append(f"Документ о качестве №{row['Номер документа']} не указан в АООК.")
            row['Статус проверки'] = STATUS['АОРПИ !=']
        elif row['_Дата отгрузки'] != row['_Качества Дата из АООК']:
            row['Акт проверки'].append(
                f"Дата {row['Дата отгрузки']} документа о качестве №{row['Номер документа']} не равна дате {row['Качества Дата из АООК']} документа о качестве в АООК.")
            row['Статус проверки'] = STATUS['АООК != дата']
            row['Расхождения'].append({'Тип': 'Дата отгрузки',
                                       'Рез': row['Качества Дата из АООК'],
                                       'Преф': 'Дата из АООК/АОСР: \n'})
        elif row['_Дата отгрузки'] > row['_Дата АООК']:
            row['Статус проверки'] = STATUS['АООК < АоРПИ']
            row['Расхождения'].append({'Тип': 'Дата отгрузки',
                                       'Рез': f"{row['Дата отгрузки']} > {row['Дата АООК']}",
                                       'Преф': ''})
        else:
            row['Статус проверки'] = STATUS['Совпало']

        if row['Номер документа'] != '':
            row['Номер документа'] = create_hyperlink(row["ПутьФайла"], row['Файл'], row['Номер документа'])
        if row['Качества Номер из АООК'] != '':
            row['Качества Номер из АООК'] = create_hyperlink(row["Путь файла АООК"], '', row['Качества Номер из АООК'])

    logging.info(f"Сверка документов ДК с АООК произведена.")
    return data_itog
