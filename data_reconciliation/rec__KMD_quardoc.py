import copy
import logging

from def_folder.data_normalization import append_value as ap, create_hyperlink

STATUS = {'Полное совпадение': 'Полное совпадение',
          'Количество': 'Не соответствует количество',
          'Наименование': 'Не соответствует наименование',
          'Марка': 'Не найдена марка',
          'Вес': 'Не соответствует Общий вес'}  # Итоговые статусы проверки


def reconciliation_data(kmd, qualdoc, aook):
    if not kmd:
        logging.warning(f"Сверка ДК с КМД не проведена! Данные по КМД отсутствуют!")
        return None
    result = copy.deepcopy(kmd)

    if not qualdoc:
        logging.warning(f"Сверка ДК с КМД не проведена! Данные по ДК отсутствуют!")
        return None

    if not aook:
        logging.warning(f"Сверка ДК с КМД не проведена! Данные по АООК отсутствуют!")
        return None

    # Сопаставление данных АООК с Качества
    for row_quardoc in qualdoc:
        row_quardoc['Сверка Качества и АООК'] = False
        if row_quardoc.get('Номер') != '':
            for row_aook in aook:
                if row_quardoc['Номер документа'] == row_aook.get('Качества Номер из АООК', ""):
                    row_quardoc['Сверка Качества и АООК'] = True
                    row_quardoc['Качества Дата из АООК'] = row_aook.get('Качества Дата из АООК', "")
                    break

    # Сопаставление данных КМД с качество с изделиями
    for row_kmd in result:
        row_kmd['Марка качества'] = []
        row_kmd['Наим качества'] = []
        for i, row_quardoc in enumerate(qualdoc):
            if row_kmd['Марка'] == row_quardoc['Марка']:  # Проверка на совпадение марки
                row_kmd['Марка качества'].append(i)
                # Проверка на совпадение наименования после марки
                if row_quardoc['Наименование'] + ' ' in row_kmd['Наименование'] or \
                        row_kmd['Наименование'] == row_quardoc['Наименование']:
                    row_kmd['Наим качества'].append(i)

    # Сверяем данные КМД с Качества
    for row_kmd in result:
        row_kmd['Расхождения'] = []
        row_kmd['Статус проверки'] = ''
        row_kmd['Акт проверки'] = []
        if row_kmd['Марка качества']:
            row_quardoc = qualdoc[row_kmd['Марка качества'][0]]  # Строка в таблице Качества

            row_quardoc["Номер документа"] = ",".join(set(row_quardoc["Номер документа"].split(', ')))
            # Наименование
            if row_kmd['Наименование'] != row_quardoc['Наименование']:
                row_kmd['Статус проверки'] = ap(row_kmd['Статус проверки'], STATUS['Наименование'])
                row_kmd['Акт проверки'].append(
                    f'Наименование элемента марки "{row_kmd["Марка"]}" в КМД не соответствует наименованию элемента в КМД №"{row_quardoc["Номер документа"]}"')
                row_kmd['Расхождения'].append({'Тип': 'Наименование',
                                               'Рез': row_quardoc['Наименование'],
                                               'Преф': 'Наименование из качества: \n'})

            # Количество
            if row_kmd['Количество'] != row_quardoc['Количество']:
                row_kmd['Статус проверки'] = ap(row_kmd['Статус проверки'], STATUS['Количество'])
                row_kmd['Акт проверки'].append(
                    f'Количество элементов марки "{row_kmd["Марка"]}", указанное в КМД, не соответствует количеству элементов, указанному в документе о качестве №"{row_quardoc["Номер документа"]}".')
                row_kmd['Расхождения'].append({'Тип': 'Количество',
                                               'Рез': row_quardoc['Количество'],
                                               'Преф': 'Количество из качества: \n'})

            # Вес общий
            if row_kmd['Вес общий'] != row_quardoc['Вес общий']:
                row_kmd['Статус проверки'] = ap(row_kmd['Статус проверки'], STATUS['Вес'])
                row_kmd['Акт проверки'].append(
                    f'Общий вес элементов марки "{row_kmd["Марка"]}", указанный в КМД, не соответствует общему весу, указанному в документе о качестве №"{row_quardoc["Номер документа"]}".')
                row_kmd['Расхождения'].append({'Тип': 'Вес общий',
                                               'Рез': str(row_quardoc['Вес общий']),
                                               'Преф': 'Вес общий из качества: \n'})
            row_kmd['Номер документа'] = create_hyperlink(row_quardoc["ПутьФайла|"], row_quardoc['Файл'], row_quardoc['Номер документа'])
        else:
            row_kmd['Статус проверки'] = STATUS['Марка']

        if row_kmd['Статус проверки'] == '':
            row_kmd['Статус проверки'] = STATUS['Полное совпадение']

    logging.info('Сверка Качество с КМД произведена.')
    return result
