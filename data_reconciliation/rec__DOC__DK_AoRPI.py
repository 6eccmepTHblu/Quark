import pandas as pd

STATUS = {'АОРПИ !=': 'Не приложен ДК',
          'АООК !=': 'ДК не указан в АоРПИ',
          'АООК != дата': 'Дата ДК не равна дате ДК в АоРПИ',
          'Совпало': 'Полное совпадение',
          'АООК < АоРПИ': 'Дата АООК раньше АоРПИ'}  # Итоговые статусы проверки


def reconciliation_data(dk, aorpi):
    if not dk:
        print('!!! Данны по ДК отсутствуют!')
        return None

    if not aorpi:
        print('!!! Данны по АоРПИ отсутствуют!')
        return None

    # Объединение по ключу
    data_itog = pd.merge(
        pd.DataFrame(dk),
        pd.DataFrame(aorpi),
        left_on='Номер документа',
        right_on='Номер док.',
        how='outer',
        suffixes=('_1', '_2'),
        validate="many_to_many"
    )
    data_itog = data_itog.drop_duplicates(subset=['Номер документа',
                                                  'Номер док.',
                                                  'Дата отгрузки',
                                                  'Дата док.'])

    # Преобразуем даты в даты
    data_itog['_Дата док.'] = pd.to_datetime(data_itog['Дата док.'], format='%d.%m.%Y', errors='coerce')
    data_itog['_Дата отгрузки'] = pd.to_datetime(data_itog['Дата отгрузки'], format='%d.%m.%Y',
                                                         errors='coerce')

    # Конвертируем данные в лист словарей
    data_itog['Расхождения'] = data_itog.apply(lambda row: [], axis=1)
    data_itog = data_itog.map(lambda x: '' if pd.isna(x) else x)
    data_itog = data_itog.to_dict(orient='records')

    # Сверяем даты между собой
    for row in data_itog:
        if row['Номер док.'] == '':
            row['Статус проверки'] = STATUS['АООК !=']
        elif row['Номер документа'] == '':
            row['Статус проверки'] = STATUS['АОРПИ !=']
        elif row['_Дата док.'] != row['_Дата отгрузки']:
            row['Статус проверки'] = STATUS['АООК != дата']
            row['Расхождения'].append({'Тип': 'Дата док.',
                                       'Рез': row['Дата отгрузки'],
                                       'Преф': 'Дата из АООК: \n'})
        else:
            row['Статус проверки'] = STATUS['Совпало']

    return data_itog
