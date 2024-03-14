import logging
import pandas as pd
import copy

from def_folder.data_normalization import append_value as ap, create_hyperlink

STATUS = {'Полное совпадение': 'Полное совпадение',
          'Количество': 'Не соответствует количество',
          'Наименование': 'Не соответствует наименование',
          'Марка': 'Не найдена марка'}  # Итоговые статусы проверки


def reconciliation_data(kmd, aorpi, aook):
    if not kmd:
        logging.warning(f"Сверка КМД с АоРПИ не проведена! Данные по КМД отсутствуют!")
        return None
    result = copy.deepcopy(kmd)

    if not aorpi:
        logging.warning(f"Сверка КМД с АоРПИ не проведена! Данные по АоРПИ отсутствуют!")
        return None

    if not aook:
        logging.warning(f"Сверка КМД с АоРПИ не проведена! Данные по АООК отсутствуют!")
        return None

    # Скрещивание строк по марке АоРПИ
    df = pd.DataFrame(aorpi)
    df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(0)
    df['Количество'] = df['Количество'].astype(int)
    aorpi = df.groupby(['Марка', 'НаимПродукции']).agg({
        'Количество': 'sum',
        'Номер': lambda x: ', '.join(x.astype(str)),
        **{col: 'first' for col in df.columns if
           col not in ['Марка', 'НаимПродукции', 'Количество', 'Номер']}
    }).reset_index()
    aorpi = aorpi.to_dict(orient='records')

    # Сопаставление данных АООК с АоРПИ
    for row_aorpi in aorpi:
        row_aorpi['Сверка АоРПИ и АООК'] = False
        if row_aorpi.get('Номер') != '':
            for row_aook in aook:
                if row_aorpi['Номер'] == row_aook.get('АоРПИ Номер из АООК', ""):
                    row_aorpi['Сверка АоРПИ и АООК'] = True
                    row_aorpi['АоРПИ Дата из АООК'] = row_aook.get('АоРПИ Дата из АООК', "")
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
        row_kmd['Акт проверки'] = []
        if row_kmd['Марка АоРПИ']:
            row_aorpi = aorpi[row_kmd['Марка АоРПИ'][0]]  # Строка в таблице АоРПИ

            row_aorpi["Номер"] = ",".join(set(row_aorpi["Номер"].split(', ')))

            # Наименование
            if row_kmd['Наименование'] != row_aorpi['НаимПродукции']:
                row_kmd['Статус проверки'] = ap(row_kmd['Статус проверки'], STATUS['Наименование'])
                row_kmd['Акт проверки'].append(
                    f'Наименование элемента марки "{row_kmd["Марка"]}" в КМД не соответствует наименованию элемента в АоРПИ №"{row_aorpi["Номер"]}"')
                row_kmd['Расхождения'].append({'Тип': 'Наименование',
                                               'Рез': row_aorpi['НаимПродукции'],
                                               'Преф': 'Наименование из АоРПИ: \n'})

            # Количество
            if str(row_kmd['Количество']) != str(row_aorpi['Количество']):
                row_kmd['Статус проверки'] = ap(row_kmd['Статус проверки'], STATUS['Количество'])
                row_kmd['Акт проверки'].append(
                    f'Количество элементов марки "{row_kmd["Марка"]}" в КМД не соответствует количеству элементов в АоРПИ №"{row_aorpi["Номер"]}".')
                row_kmd['Расхождения'].append({'Тип': 'Количество',
                                               'Рез': str(row_aorpi['Количество']),
                                               'Преф': 'Количество из АоРПИ: \n'})

            if len(row_aorpi['Номер']) < 250:
                row_kmd['Номер АоРПИ'] = create_hyperlink(row_aorpi["ПутьФайла|"], row_aorpi['Файл'],
                                                          row_aorpi['Номер'])
            else:
                row_kmd['Номер АоРПИ'] = row_aorpi['Номер']

        else:
            row_kmd['Статус проверки'] = STATUS['Марка']
        if row_kmd['Статус проверки'] == '':
            row_kmd['Статус проверки'] = STATUS['Полное совпадение']

    logging.info('Сверка АоРПИ с КМД произведена.')
    return result
