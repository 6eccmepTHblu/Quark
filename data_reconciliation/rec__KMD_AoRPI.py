from def_folder.data_normalization import append_value as ap
import copy


STATUS = {'Полное совпадение': 'Полное совпадение',
          'Количество': 'Количество',
          'Наименование': 'Наименование',
          'Марка': 'Не найдена марка'}  # Итоговые статусы проверки


def reconciliation_data(kmd, aorpi, aook):
    if not kmd:
        print('!!! Данны по КМД отсутствуют!')
        return None
    result = copy.deepcopy(kmd)

    if not aorpi:
        print('!!! Данны по АоРПИ отсутствуют!')
        return None

    if not aook:
        print('!!! Данны по АООК отсутствуют!')
        return None

    # Сопаставление данных АООК с АоРПИ
    for row_aorpi in aorpi:
        row_aorpi['Сверка АоРПИ и АООК'] = False
        if row_aorpi.get('Номер') != '':
            for row_aook in aook:
                if row_aorpi['Номер'] == row_aook['АоРПИ Номер из АООК']:
                    row_aorpi['Сверка АоРПИ и АООК'] = True
                    row_aorpi['АоРПИ Дата из АООК'] = row_aook['АоРПИ Дата из АООК']
                    break

    # Сопаставление данных КМД с АоРПИ
    for row_kmd in result:
        row_kmd['Марка АоРПИ'] = []
        row_kmd['Наим АоРПИ'] = []
        for i, row_aorpi in enumerate(aorpi):
            if row_kmd['Марка'] == row_aorpi['Марка']:  # Проверка на совпадение марки
                row_kmd['Марка АоРПИ'].append(i)
                # Проверка на совпадение наименования после марки
                if row_aorpi['НаимПродукции'] + ' ' in row_kmd['Наименование'] or \
                        row_kmd['Наименование'] == row_aorpi['НаимПродукции']:
                    row_kmd['Наим АоРПИ'].append(i)

    # Сверяем данные КМД с АоРПИ
    for row_kmd in result:
        row_kmd['Расхождения'] = []
        row_kmd['Статус проверки'] = ''
        if row_kmd['Марка АоРПИ']:
            row_aorpi = aorpi[row_kmd['Марка АоРПИ'][0]]  # Строка в таблице АоРПИ
            row_kmd['Номер АоРПИ'] = row_aorpi['Номер']

            # Наименование
            if row_kmd['Наименование'] != row_aorpi['НаимПродукции']:
                row_kmd['Статус проверки'] = ap(row_kmd['Статус проверки'], STATUS['Наименование'])
                row_kmd['Расхождения'].append({'Тип': 'Наименование',
                                               'Рез': row_aorpi['НаимПродукции'],
                                               'Преф': 'Наименование из АоРПИ: \n'})

            # Количество
            if row_kmd['Количество'] != row_aorpi['Количество']:
                row_kmd['Статус проверки'] = ap(row_kmd['Статус проверки'], STATUS['Количество'])
                row_kmd['Расхождения'].append({'Тип': 'Количество',
                                               'Рез': row_aorpi['Количество'],
                                               'Преф': 'Количество из АоРПИ: \n'})
        else:
            row_kmd['Статус проверки'] = STATUS['Марка']
        if row_kmd['Статус проверки'] == '':
            row_kmd['Статус проверки'] = STATUS['Полное совпадение']

    print('Сверка АоРПИ с КМД произведена.')
    return result
