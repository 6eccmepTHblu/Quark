from def_folder import data_collection as coll
from def_folder.data_normalization import append_value as ap

STATUS = {'Полное совпадение': 'Полное совпадение',
          'Количество': 'Количество',
          'Наименование': 'Наименование',
          'Марка': 'Не найдена марка',
          'Дата из АоРПИ': 'Дата ДК из АоРПИ',
          'Дата ДК и АоРПИ': 'Дата ДК позже АоРПИ',
          'Номер из АоРПИ': 'Дата ДК не найдена в АОРПИ',
          'АоРПИ номер в ДК': 'Номер ДК не найден в АОРПИ',}  # Итоговые статусы проверки


def reconciliation_data(quardoc: list, aorpi_izdel: list, aorpi_docum: list) -> list:
    if not aorpi_izdel:
        print('!!! Данны по АоРПИ изделя отсутствуют!')
        return []

    if not quardoc:
        print('!!! Данны по Качеству отсутствуют!')
        return []

    if not aorpi_docum:
        print('!!! Данны по АоРПИ документы отсутствуют!')
        return []

    # Сопаставление данных качество с документами
    for row_quardoc in quardoc:
        row_quardoc['Сверка кач. докум.'] = []
        for i, row_docum in enumerate(aorpi_docum):
            if row_quardoc.get('Номер документа', '+') == row_docum.get('Номер док.', '-'):
                row_quardoc['Сверка кач. докум.'].append(i)
                row_quardoc['Номер АоРПИ'] = row_docum['Номер']
                row_quardoc['Дата АоРПИ'] = row_docum['Дата']
                row_quardoc['Дата АоРПИ Отгрузки'] = row_docum['Дата док.']

    # Сопаставление данных качество с изделиями
    for row_quardoc in quardoc:
        row_quardoc['Марка АоРПИ'] = []
        row_quardoc['Наим АоРПИ'] = []
        for i, row_izdel in enumerate(aorpi_izdel):
            if row_quardoc['Марка'] == row_izdel['Марка']:  # Проверка на совпадение марки
                row_quardoc['Марка АоРПИ'].append(i)
                # Проверка на совпадение наименования после марки
                if row_izdel['НаимПродукции'] + ' ' in row_quardoc['Наименование'] or \
                        row_quardoc['Наименование'] == row_izdel['НаимПродукции']:
                    row_quardoc['Наим АоРПИ'].append(i)
                    break

    # Сверяем данные КМД с АоРПИ
    for row_quardoc in quardoc:
        row_quardoc.setdefault('Дата отгрузки','01.01.1990')
        row_quardoc['Расхождения'] = []
        row_quardoc['Статус проверки'] = ''
        if row_quardoc['Марка АоРПИ']:
            row_izdel = aorpi_izdel[row_quardoc['Марка АоРПИ'][0]]  # Строка в таблице АоРПИ
            row_quardoc['Номер АоРПИ'] = row_izdel['Номер']
            row_quardoc['Дата АоРПИ'] = row_izdel['Дата']

            # Наименование
            if row_quardoc['Наименование'] != row_izdel['НаимПродукции']:
                row_quardoc['Статус проверки'] = ap(row_quardoc['Статус проверки'], STATUS['Наименование'])
                row_quardoc['Расхождения'].append({'Тип': 'Наименование',
                                               'Рез': row_izdel['НаимПродукции'],
                                               'Преф': 'Наименование из АоРПИ: \n'})

            # Количество
            if row_quardoc['Количество'] != row_izdel['Количество']:
                row_quardoc['Статус проверки'] = ap(row_quardoc['Статус проверки'], STATUS['Количество'])
                row_quardoc['Расхождения'].append({'Тип': 'Количество',
                                               'Рез': row_izdel['Количество'],
                                               'Преф': 'Количество из АоРПИ: \n'})

            # Дата Качества самом в АоРПИ
            if len(row_quardoc['Сверка кач. докум.']) != 0:
                if row_quardoc['Дата отгрузки'] != row_quardoc.setdefault('Дата АоРПИ Отгрузки', ''):
                    row_quardoc['Статус проверки'] = ap(row_quardoc['Статус проверки'], STATUS['Дата из АоРПИ'])
                    row_quardoc['Расхождения'].append({'Тип': 'Дата отгрузки',
                                                   'Рез': row_quardoc['Дата АоРПИ Отгрузки'],
                                                   'Преф': 'Дата ДК из АоРПИ: \n'})

                # Дата ДК позже АоРПИ
                elif coll.checks_dates(row_quardoc.setdefault('Дата отгрузки', ''),
                                       row_quardoc.setdefault('Дата АоРПИ', row_quardoc['Дата отгрузки']), '>'):
                    row_quardoc['Статус проверки'] = ap(row_quardoc['Статус проверки'], STATUS['Дата ДК и АоРПИ'])
                    row_quardoc['Расхождения'].append({'Тип': 'Дата АоРПИ',
                                                   'Рез': row_quardoc['Дата АоРПИ'],
                                                   'Преф': 'Дата АоРПИ: \n'})
            else:
                row_quardoc['Статус проверки'] = ap(row_quardoc['Статус проверки'], STATUS['АоРПИ номер в ДК'])

        else:
            row_quardoc['Статус проверки'] = STATUS['Марка']

        if row_quardoc['Статус проверки'] == '':
            row_quardoc['Статус проверки'] = STATUS['Полное совпадение']

    print('Сверка Качества с АоРПИ произведена.')
    return quardoc
