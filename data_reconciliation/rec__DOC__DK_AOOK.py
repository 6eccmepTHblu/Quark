import pandas as pd

from def_folder.data_normalization import create_hyperlink

STATUS = {'АОРПИ !=': 'ДК не указан в АООК',
          'АООК !=': 'Не приложен ДК',
          'АООК != дата': 'Дата ДК не равна дате ДК в АООК',
          'Совпало': 'Полное совпадение',
          'АООК < АоРПИ': 'Дата АООК раньше АоРПИ'}  # Итоговые статусы проверки


def reconciliation_data(dk, aook, date_aook):
    if not dk:
        print('!!! Данны по ДК отсутствуют!')
        return None

    if not aook:
        print('!!! Данны по АООК отсутствуют!')
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
        if row['Номер документа'] != '':
            row['Номер документа'] = create_hyperlink(row["ПутьФайла"], row['Файл'], row['Номер документа'])
        if row['Качества Номер из АООК'] != '':
            row['Качества Номер из АООК'] = create_hyperlink(row["Путь файла АООК"], '', row['Качества Номер из АООК'])

        if row['Номер документа'] == '':
            row['Статус проверки'] = STATUS['АООК !=']
        elif row['Качества Номер из АООК'] == '':
            row['Статус проверки'] = STATUS['АОРПИ !=']
        elif row['_Дата отгрузки'] != row['_Качества Дата из АООК']:
            row['Статус проверки'] = STATUS['АООК != дата']
            row['Расхождения'].append({'Тип': 'Дата отгрузки',
                                       'Рез': row['Качества Дата из АООК'],
                                       'Преф': 'Дата из АООК: \n'})
        elif row['_Дата отгрузки'] > row['_Дата АООК']:
            row['Статус проверки'] = STATUS['АООК < АоРПИ']
            row['Расхождения'].append({'Тип': 'Дата отгрузки',
                                       'Рез': f"{row['Дата отгрузки']} > {row['Дата АООК']}",
                                       'Преф': ''})
        else:
            row['Статус проверки'] = STATUS['Совпало']

    return data_itog
