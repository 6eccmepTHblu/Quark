import logging
import traceback
import os
import subprocess

from sys import argv, exit
from data_output.data_in_excel import main as xl
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
    path = r'D:\Работа\Силенко Д.Т\Задача 2(Путхон - Кварк V2)\Данные 2'

FAIL_NAME = os.path.join(path, "Акт промежуточной проверки.xlsx")


def main():
    # Удаляем файл, если он етсь
    if os.path.exists(FAIL_NAME):
        try:
            os.remove(FAIL_NAME)
        except:
            logging.warning(f"'Акт промежуточной проверки.xlsx' открыт, закройте файл и начните снова!")
            return

    # Загрузка данных для нормирования заголовков
    fun.load_data()

    # Собираем данные
    all_data = {
        'АОРПИ': fun.collects_data_by_type(path, ['АОРПИ Усть-Луга'], '-'),
        'АОРПИ изделия': fun.collects_data_by_type(path, ['АОРПИ Усть-Луга'], 'АОРПИ Усть-Луга изделия'),
        'АОРПИ докум': fun.collects_data_by_type(path, ['АОРПИ Усть-Луга'], 'АОРПИ Усть-Луга сопр.док.'),
        'АООК дата': excel_AOOK.get_date(path),
        'АООК номер': excel_AOOK.get_number(path),
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
            all_data['ДК реестр'],
            all_data['АОРПИ докум']
        ),
        'Докум->КМД=АООК': rec__DOC__DK_AOOK.reconciliation_data(
            all_data['ДК реестр'],
            all_data['АООК качество'],
            all_data['АООК дата']
        ),
        'Докум->АоРПИ=АООК': rec__DOC__AoRPI_AOOK.reconciliation_data(
            all_data['АОРПИ'],
            all_data['АООК ПКМ'],
            all_data['АООК дата']
        ),
        'АоРПИ=КМД': rec__KMD_AoRPI.reconciliation_data(
            all_data['КМД'],
            all_data['АОРПИ изделия'],
            all_data['АООК ПКМ']
        ),
        'Заключения=ЖСР': rec__JSR_CSV.reconciliation_data(
            all_data['ЖСР'],
            all_data['Заключения']
        ),
        'Качество=АоРПИ': rec__quardoc_AoRPI.reconciliation_data(
            all_data['ДК'],
            all_data['АОРПИ изделия'],
            all_data['АОРПИ докум']
        ),
        'Качество=КМД': rec__KMD_quardoc.reconciliation_data(
            all_data['КМД'],
            all_data['ДК'],
            all_data['АООК качество']
        )
    }

    # Вывод в Excel
    out__KMD_AoRPI.data_output(all_result_check["АоРПИ=КМД"], FAIL_NAME)
    out__JSR_CSV.data_output(all_result_check["Заключения=ЖСР"], FAIL_NAME)
    out__quardoc_AoRPI.data_output(all_result_check["Качество=АоРПИ"], FAIL_NAME)
    out__KMD_qualdoc.data_output(all_result_check["Качество=КМД"], FAIL_NAME)
    out__DOC__AoRPI_AOOK.data_output(all_result_check["Докум->АоРПИ=АООК"], FAIL_NAME, all_data['АООК дата'])
    out__DOC__DK_AOOK.data_output(all_result_check["Докум->КМД=АООК"], FAIL_NAME, all_data['АООК дата'])
    out__DOC__DK_AoRPI.data_output(all_result_check["Докум->КМД=АоРПИ"], FAIL_NAME, all_data['АООК дата'])

    # Добовляем данные в книгу
    xl(FAIL_NAME, 'АОРПИ изделия', all_data['АОРПИ изделия'])
    xl(FAIL_NAME, 'АОРПИ', all_data['АОРПИ'])
    xl(FAIL_NAME, 'АОРПИ докум', all_data['АОРПИ докум'])
    xl(FAIL_NAME, 'АООК ПКМ', all_data['АООК ПКМ'])
    xl(FAIL_NAME, 'КМД', all_data['КМД'])
    xl(FAIL_NAME, 'Заключения', all_data['Заключения'])
    xl(FAIL_NAME, 'ЖСР', all_data['ЖСР'])
    xl(FAIL_NAME, 'ДК', all_data['ДК'])
    xl(FAIL_NAME, 'ДК реестр', all_data['ДК реестр'])

    # Создаем акт проверки
    data_reports = create_reports(all_result_check)
    create_act_checking(FAIL_NAME,
                        all_data['АООК номер'],
                        data_reports)


if __name__ == "__main__":
    try:
        file_log = logging.FileHandler(path + '\\Лог работы скрипта.txt', mode='w')
        console_out = logging.StreamHandler()
        logging.basicConfig(handlers=(file_log, console_out), level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
        logging.info(f"Путь реестра - '{path}'")
        main()
    except Exception as ex:
        # Настройка конфигурации логирования
        print('Ошибка:\n', traceback.format_exc())
        logging.error("====================================================================")
        logging.exception("Произошла ошибка: %s", str(ex))
        exit(1)
    subprocess.Popen(['notepad.exe', path + '\\Лог работы скрипта.txt'])
