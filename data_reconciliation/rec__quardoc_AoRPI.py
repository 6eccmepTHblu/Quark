import pandas as pd
import os
from def_folder.data_normalization import (append_value as ap,
                                           csv_to_pdf_path)

STATUS = {'Полное совпадение': 'Полное совпадение',
          'Количество': 'Количество',
          'Наименование': 'Наименование',
          'Марка': 'Не найдена марка'}  # Итоговые статусы проверки


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

    # Скрещивание строк по марке
    df = pd.DataFrame(quardoc)
    df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce')
    quardoc = df.groupby(['Марка', 'Наименование']).agg({
        'Количество': 'sum',
        'Номер документа': lambda x: ', '.join(x.astype(str).unique()),
        **{col: 'first' for col in df.columns if
           col not in ['Марка', 'Наименование', 'Количество', 'Номер АоРПИ', 'Номер документа']}
    }).reset_index()
    quardoc = quardoc.to_dict(orient='records')

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
        if row_quardoc['Марка АоРПИ']:
            row_izdel = aorpi_izdel[row_quardoc['Марка АоРПИ'][0]]  # Строка в таблице АоРПИ
            row_quardoc['Номер АоРПИ'] = row_izdel['Номер']
            row_quardoc['Дата АоРПИ'] = row_izdel['Дата']
            row_quardoc[
                'Номер АоРПИ'] = f'=HYPERLINK("{csv_to_pdf_path(row_izdel["ПутьФайла|"], row_izdel["Файл"])}", "{row_quardoc["Номер АоРПИ"]}")'

            # Наименование
            if row_quardoc['Наименование'] != row_izdel['НаимПродукции']:
                row_quardoc['Статус проверки'] = ap(row_quardoc['Статус проверки'], STATUS['Наименование'])
                row_quardoc['Расхождения'].append({'Тип': 'Наименование',
                                                   'Рез': row_izdel['НаимПродукции'],
                                                   'Преф': 'Наименование из АоРПИ: \n'})

            # Количество
            if str(row_quardoc['Количество']) != str(row_izdel['Количество']):
                row_quardoc['Статус проверки'] = ap(row_quardoc['Статус проверки'], STATUS['Количество'])
                row_quardoc['Расхождения'].append({'Тип': 'Количество',
                                                   'Рез': row_izdel['Количество'],
                                                   'Преф': 'Количество из АоРПИ: \n'})
        else:
            row_quardoc['Статус проверки'] = STATUS['Марка']

        if row_quardoc['Статус проверки'] == '':
            row_quardoc['Статус проверки'] = STATUS['Полное совпадение']

    print('Сверка Качества с АоРПИ произведена.')
    return quardoc
