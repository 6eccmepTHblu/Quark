import logging
import pandas as pd
import os

from def_folder.data_normalization import (append_value as ap,
                                           create_hyperlink)

STATUS = {'Полное совпадение': 'Полное совпадение',
          'Количество': 'Не соответствует количество',
          'Наименование': 'Не соответствует наименование',
          'Марка': 'Не найдена марка'}  # Итоговые статусы проверки


def reconciliation_data(quardoc: list, aorpi_izdel: list, aorpi_docum: list) -> list:
    if not aorpi_izdel:
        logging.warning(f"Сверка ДК с АоРПИ не проведена! Данные по АоРПИ изделия отсутствуют!")
        return []

    if not quardoc:
        logging.warning(f"Сверка ДК с АоРПИ не проведена! Данные по ДК отсутствуют!")
        return []

    if not aorpi_docum:
        logging.warning(f"Сверка ДК с АоРПИ не проведена! Данные по АоРПИ документы отсутствуют!")
        return []

    # Скрещивание строк по марке ДК
    df = pd.DataFrame(quardoc)
    df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(0)
    df['Количество'] = df['Количество'].astype(int)
    quardoc = df.groupby(['Марка', 'Наименование']).agg({
        'Количество': 'sum',
        'Номер документа': lambda x: ', '.join(x.astype(str)),
        **{col: 'first' for col in df.columns if
           col not in ['Марка', 'Наименование', 'Количество', 'Номер АоРПИ', 'Номер документа']}
    }).reset_index()
    quardoc = quardoc.to_dict(orient='records')

    # Скрещивание строк по марке АоРПИ
    df = pd.DataFrame(aorpi_izdel)
    df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(0)
    df['Количество'] = df['Количество'].astype(int)
    aorpi_izdel = df.groupby(['Марка', 'НаимПродукции']).agg({
        'Количество': 'sum',
        'Номер': lambda x: ', '.join(x.astype(str)),
        **{col: 'first' for col in df.columns if
           col not in ['Марка', 'НаимПродукции', 'Количество', 'Номер']}
    }).reset_index()
    aorpi_izdel = aorpi_izdel.to_dict(orient='records')

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
        row_quardoc.setdefault('Дата отгрузки', '01.01.1990')
        row_quardoc['Расхождения'] = []
        row_quardoc['Статус проверки'] = ''
        row_quardoc['Акт проверки'] = []
        row_quardoc["Номер документа"] = ",".join(set(row_quardoc["Номер документа"].split(', ')))
        if row_quardoc['Марка АоРПИ']:
            row_izdel = aorpi_izdel[row_quardoc['Марка АоРПИ'][0]]  # Строка в таблице АоРПИ
            row_quardoc['Дата АоРПИ'] = row_izdel['Дата']

            row_izdel["Номер"] = ",".join(set(row_izdel["Номер"].split(', ')))

            # Наименование
            if row_quardoc['Наименование'] != row_izdel['НаимПродукции']:
                row_quardoc['Статус проверки'] = ap(row_quardoc['Статус проверки'], STATUS['Наименование'])
                row_quardoc['Акт проверки'].append(
                    f'Наименование элемента марки "{row_quardoc["Марка"]}" в КМД не соответствует наименованию элемента в АоРПИ №"{row_izdel["Номер"]}"')
                row_quardoc['Расхождения'].append({'Тип': 'Наименование',
                                                   'Рез': row_izdel['НаимПродукции'],
                                                   'Преф': 'Наименование из АоРПИ: \n'})

            # Количество
            if str(row_quardoc['Количество']) != str(row_izdel['Количество']):
                row_quardoc['Статус проверки'] = ap(row_quardoc['Статус проверки'], STATUS['Количество'])
                row_quardoc['Акт проверки'].append(
                    f'Количество элементов марки "{row_quardoc["Марка"]}" в КМД не соответствует количеству элементов в АоРПИ №"{row_izdel["Номер"]}".')
                row_quardoc['Расхождения'].append({'Тип': 'Количество',
                                                   'Рез': row_izdel['Количество'],
                                                   'Преф': 'Количество из АоРПИ: \n'})

            if len(row_izdel['Номер']) < 250:
                row_quardoc['Номер АоРПИ'] = create_hyperlink(row_izdel["ПутьФайла|"], row_izdel['Файл'], row_izdel['Номер'])
            else:
                row_quardoc['Номер АоРПИ'] = row_izdel['Номер']
        else:
            row_quardoc['Статус проверки'] = STATUS['Марка']
            row_quardoc['Акт проверки'].append(
                f'Указанная в документе о качестве №"{row_quardoc["Номер документа"]}" марка элемента "{row_quardoc["Марка"]}" не найдена в АоРПИ.')


        if len(row_quardoc['Номер документа']) < 250:
            row_quardoc['Номер документа'] = create_hyperlink(row_quardoc["ПутьФайла|"], row_quardoc['Файл'], row_quardoc['Номер документа'])
        else:
            row_quardoc['Номер документа'] = row_quardoc['Номер документа']

        if row_quardoc['Статус проверки'] == '':
            row_quardoc['Статус проверки'] = STATUS['Полное совпадение']

    logging.info(f"Сверка Качества с АоРПИ произведена.")
    return quardoc
