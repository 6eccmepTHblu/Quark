import pprint
import numpy as np

import pandas as pd

STATUS = {'АОРПИ !=': 'Не приложен АОРПИ',
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

    # Объединение по ключу
    data_itog = pd.merge(
        pd.DataFrame(aorpi),
        pd.DataFrame(aook),
        left_on='Номер',
        right_on='АоРПИ Номер из АООК',
        how='outer',
        suffixes=('_1', '_2'),
        validate="many_to_many"
    )

    # Проверка отсутствия 'Номер' в data_itog
    condition_no_aorpi = data_itog['Номер'].isnull()
    data_itog.loc[condition_no_aorpi, 'Статус проверки'] = STATUS['АООК !=']

    # Проверка отсутствия 'АоРПИ Номер из АООК' в data_itog
    condition_no_aook = data_itog['АоРПИ Номер из АООК'].isnull()
    data_itog.loc[condition_no_aook, 'Статус проверки'] = STATUS['АОРПИ !=']

    # Проверка дат АоРПИ и АООК
    data_itog['Дата'] = pd.to_datetime(data_itog['Дата'], format='%d.%m.%Y', errors='coerce')
    data_itog['АоРПИ Дата из АООК'] = pd.to_datetime(data_itog['АоРПИ Дата из АООК'], format='%d.%m.%Y',
                                                     errors='coerce')
    data_itog.replace('nan', '', inplace=True)
    condition = data_itog['Статус проверки'].eq("")
    data_itog.loc[condition, 'Статус проверки'] = np.where(
        data_itog.loc[condition, 'Дата'] == data_itog.loc[condition, 'АоРПИ Дата из АООК'],
        '',
        STATUS['АООК != дата']
    )
    # Конвертируем данные в лист словарей
    data_itog['Дата'] = data_itog['Дата'].dt.strftime('%d.%m.%Y')
    data_itog['АоРПИ Дата из АООК'] = data_itog['АоРПИ Дата из АООК'].dt.strftime('%d.%m.%Y')
    data_itog = data_itog.map(lambda x: '' if pd.isna(x) else x)
    data_itog = data_itog.to_dict(orient='records')

    return data_itog
