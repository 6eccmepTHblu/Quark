import logging
import traceback
import os

from sys import argv, exit

from aсt_checking import create_act_checking, create_reports
from def_folder import data_collection as fun
from loading_data import excel_KMD, excel_AOOK, excel_JSR
from data_reconciliation import (
    rec__KMD_AoRPI,
    rec__JSR_CSV,
    rec__quardoc_AoRPI,
    rec__KMD_quardoc,
    rec__DOC__AoRPI_AOOK,
    rec__DOC__DK_AOOK,
    rec__DOC__DK_AoRPI
)
from data_output import (
    out__KMD_AoRPI,
    out__JSR_CSV,
    out__quardoc_AoRPI,
    out__KMD_qualdoc,
    out__DOC__AoRPI_AOOK,
    out__DOC__DK_AOOK,
    out__DOC__DK_AoRPI
)

try:
    script, path = argv
except:
    path = r'D:\Работа\Силенко Д.Т\Задача 2(Путхон - Кварк V2)\Данные 4'

FAIL_NAME = os.path.join(path, "Акт промежуточной проверки.xlsx")


def main():
    # Удаляем файл, если он етсь
    if os.path.exists(FAIL_NAME):
        os.remove(FAIL_NAME)

    # Собираем данные
    all_data = {
        'АОРПИ': fun.collects_data_by_type(path, ['АОРПИ Усть-Луга'], '-'),
        'АОРПИ изделия': fun.collects_data_by_type(path, ['АОРПИ Усть-Луга'], 'АОРПИ Усть-Луга изделия'),
        'АОРПИ докум': fun.collects_data_by_type(path, ['АОРПИ Усть-Луга'], 'АОРПИ Усть-Луга сопр.док.'),
        'АООК дата': excel_AOOK.get_date(path),
        'АООК ПКМ': excel_AOOK.get_data(path, '*акт о проведении контрольного мероприятия*', 'АоРПИ'),
        'АООК качество': excel_AOOK.get_data(path, '*документ о качестве стальных*', 'Качества'),
        'КМД': excel_KMD.get_data(path),
        'Заключения': fun.collects_data_by_type(path, ['ВИК Лимак', 'ПВК Лимак']),
        'ЖСР': excel_JSR.get_data(path),
        'ДК': fun.collects_data_by_type(path, ['Документ о качестве']),
        'ДК реестр': fun.collects_data_by_type(path, ['Документ о качестве'], '-')
    }

    # Сверки
    all_result_check = {
        'Докум->КМД=АоРПИ': rec__DOC__DK_AoRPI.reconciliation_data(
            all_data['ДК реестр'][:],
            all_data['АОРПИ докум'][:]
        ),
        'Докум->КМД=АООК': rec__DOC__DK_AOOK.reconciliation_data(
            all_data['ДК реестр'][:],
            all_data['АООК качество'][:],
            all_data['АООК дата']
        ),
        'Докум->АоРПИ=АООК': rec__DOC__AoRPI_AOOK.reconciliation_data(
            all_data['АОРПИ'][:],
            all_data['АООК ПКМ'][:],
            all_data['АООК дата']
        ),
        'АоРПИ=КМД': rec__KMD_AoRPI.reconciliation_data(
            all_data['КМД'][:],
            all_data['АОРПИ изделия'][:],
            all_data['АООК ПКМ'][:]
        ),
        'Заключения=ЖСР': rec__JSR_CSV.reconciliation_data(
            all_data['ЖСР'][:],
            all_data['Заключения'][:]
        ),
        'Качество=АоРПИ': rec__quardoc_AoRPI.reconciliation_data(
            all_data['ДК'[:]],
            all_data['АОРПИ изделия'][:],
            all_data['АОРПИ докум'][:]
        ),
        'Качество=КМД': rec__KMD_quardoc.reconciliation_data(
            all_data['КМД'][:],
            all_data['ДК'][:],
            all_data['АООК качество'][:]
        )
    }

    # Вывод в Excel
    out__KMD_AoRPI.data_output(all_result_check["АоРПИ=КМД"], FAIL_NAME)
    out__JSR_CSV.data_output(all_result_check["Заключения=ЖСР"], FAIL_NAME)
    out__quardoc_AoRPI.data_output(all_result_check["Качество=АоРПИ"], FAIL_NAME)
    out__KMD_qualdoc.data_output(all_result_check["Качество=КМД"], FAIL_NAME)
    out__DOC__AoRPI_AOOK.data_output(all_result_check["Докум->АоРПИ=АООК"], FAIL_NAME)
    out__DOC__DK_AOOK.data_output(all_result_check["Докум->КМД=АООК"], FAIL_NAME)
    out__DOC__DK_AoRPI.data_output(all_result_check["Докум->КМД=АоРПИ"], FAIL_NAME)

    # Создаем акт проверки
    data_reports = create_reports(all_result_check)
    create_act_checking(FAIL_NAME,
                        "LMK/80-050/KМ1-ОК/1",
                        data_reports)


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        # Настройка конфигурации логирования
        logging.basicConfig(filename=path + '\\Отчет об ошибках.txt', level=logging.ERROR,
                            format='%(asctime)s - %(levelname)s - %(message)s')
        print('Ошибка:\n', traceback.format_exc())
        logging.exception("Произошла ошибка: %s", str(ex))
        exit(1)
