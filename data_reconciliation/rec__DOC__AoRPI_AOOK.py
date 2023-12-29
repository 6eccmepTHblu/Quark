from def_folder import data_normalization as norm
from def_folder import data_collection as coll
from def_folder.normalization import NOT_SPECIFIED

STATUS = {'АОРПИ !=': 'Не указан в АОРПИ',
          'АООК !=': 'Не указан в АООК',
          'АООК != дата': 'Не равна дате АоРПИ в АООК',
          'АООК < АоРПИ': 'Наньше чем дата АоРПИ'}  # Итоговые статусы проверки


def reconciliation_data(aorpi, aook, date_aook):
    if not aorpi:
        print('!!! Данны по АоРПИ отсутствуют!')
        return None

    if not aook:
        print('!!! Данны по АООК отсутствуют!')
        return None

    data_itog = [{'Номер АОРПИ': row['Номер'], 'Дата АОРПИ': row['Дата']} for row in aorpi]

    data_itog_dict = {row['Номер АОРПИ']: row for row in data_itog}
    for row in aook:
        if row.get('АоРПИ Номер из АООК', '+') in data_itog_dict:
            row_itog = data_itog_dict[row['АоРПИ Номер из АООК']]
            row_itog['Номер АООК'] = row['АоРПИ Номер из АООК']
            row_itog['Дата АООК'] = row['АоРПИ Дата из АООК']
        else:
            temp = {}
            temp['Номер АООК'] = row['АоРПИ Номер из АООК']
            temp['Дата АООК'] = row['АоРПИ Дата из АООК']
            data_itog.append(temp)

    for row in data_itog:
        row['Статус проверки'] = ''
        row['Расхождения'] = []
        if 'Номер АООК' in row and 'Номер АОРПИ' in row:
            if row['Дата АОРПИ'] == row['Дата АООК']:
                continue
            else:
                row['Статус проверки'] = STATUS['АООК != дата']
                row['Расхождения'].append({'Тип': 'Дата АОРПИ',
                                           'Рез': row['Дата АООК'],
                                           'Преф': 'Дата из АООК: \n'})
        elif 'Номер АООК' in row:
            row['Статус проверки'] = STATUS['АОРПИ !=']
            row['Дата АОРПИ'] = row['Дата АООК']
        elif 'Номер АОРПИ' in row:
            row['Статус проверки'] = STATUS['АООК !=']

    print('Сверка CSV(Заключений) с ЖСР произведена.')
    return data_itog
