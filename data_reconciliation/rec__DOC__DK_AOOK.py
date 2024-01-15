import pandas as pd

STATUS = {'АОРПИ !=': 'Не приложен в ДК',
          'АООК !=': 'Не указан в АООК',
          'АООК != дата': 'Не равна дате ДК в АООК',
          'АООК < АоРПИ': 'Наньше чем дата ДК'}  # Итоговые статусы проверки


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

    return data_itog
