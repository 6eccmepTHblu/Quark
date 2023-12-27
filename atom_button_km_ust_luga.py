import logging
import traceback
import os

from sys import argv, exit
from def_folder import data_collection as fun
from loading_data import excel_KMD, excel_AOOK, excel_JSR
from data_reconciliation import rec__KMD_AoRPI, rec__JSR_CSV, rec__quardoc_AoRPI, rec__KMD_quardoc
from data_output import out__KMD_AoRPI, out__JSR_CSV, out__quardoc_AoRPI, out__KMD_qualdoc
from data_output.data_in_excel import main as xl


try:
    script, path = argv
except:
    path = r'D:\Работа\Силенко Д.Т\Задача 2(Путхон - Кварк V2)\Данные 4'


FAIL_NAME = os.path.join(path, "Акт промежуточной проверки.xlsx")

# Сверка АОРПИ(+АООК) с КМД
def reconciles_documents_kmd_aorpi():
    print('\nНачата сверка КМД с АоРПИ и АООК')
    # Получаем данные заключений из АОРПИ
    data_aorpi = fun.collects_data_by_type(path, ['АОРПИ Усть-Луга'])

    # Получаем данные из КМД
    data_kmd = excel_KMD.get_data(path)

    # Получаем данные из АООК
    data_aook, date_aook = excel_AOOK.get_data(path,'*акт о проведении контрольного мероприятия*', 'АоРПИ')

    # Сверяем данные
    result = rec__KMD_AoRPI.reconciliation_data(data_kmd, data_aorpi, data_aook, date_aook)

    # Вывод данныех в Excel
    out__KMD_AoRPI.data_output(result, FAIL_NAME)

    # Добовляем данные в книгу
    # xl(FAIL_NAME, 'КМД->АоРПИ', data_kmd)
    # xl(FAIL_NAME, 'АоРПИ->КМД', data_aorpi)
    # xl(FAIL_NAME, 'АООК->АоРПИ', data_aook)


# Сверка ВИК/ПВК с ЖСР
def reconciles_documents_jsr_csv():
    print('\nНачата сверка ЖСР с НК!')
    # Получаем данные заключений из CSV
    data_csv = fun.collects_data_by_type(path, ['ВИК Лимак', 'ПВК Лимак'])

    # Получаем данные из ЖСР
    data_jsr = excel_JSR.get_data(path)

    # Сверяем данные
    result = rec__JSR_CSV.reconciliation_data(data_jsr, data_csv)

    # Вывод данныех в Excel
    out__JSR_CSV.data_output(result, FAIL_NAME)

    # Добовляем данные в книгу
    # xl(FAIL_NAME, 'ЖСР', data_jsr)
    # xl(FAIL_NAME, 'НК', data_csv)


# Сверка Качество с АоРПИ
def reconciles_documents_aorpi_qualdoc():
    print('\nНачата сверка Качества с АоРПИ')
    # Получаем данные заключений из документа о качестве
    data_qualdoc = fun.collects_data_by_type(path, ['Документ о качестве'])
    data_qualdoc = fun.sums_rows_by_keys(data_qualdoc,
                                         ['Марка', 'Наименование'],
                                         ['Количество', 'Вес общий'])

    # Получаем данные заключений из АОРПИ
    aorpi_izdel = fun.collects_data_by_type(path, ['АОРПИ Усть-Луга'], 'АОРПИ Усть-Луга изделия')
    aorpi_docum = fun.collects_data_by_type(path, ['АОРПИ Усть-Луга'], 'АОРПИ Усть-Луга сопр.док.')

    # Сверяем данные
    result = rec__quardoc_AoRPI.reconciliation_data(data_qualdoc, aorpi_izdel, aorpi_docum)

    # Вывод данныех в Excel
    out__quardoc_AoRPI.data_output(result, FAIL_NAME)

    # # Добовляем данные в книгу
    # xl(FAIL_NAME, 'Качество->АоРПИ', data_qualdoc)
    # xl(FAIL_NAME, 'Изделия', aorpi_izdel)
    # xl(FAIL_NAME, 'Сопроводит.', aorpi_docum)



# Сверка Качества(+АООК) с КМД
def reconciles_documents_kmd_qualdoc():
    print('\nНачата сверка Качества с АоРПИ и АООК')
    # Получаем данные заключений из документа о качестве
    data_qualdoc = fun.collects_data_by_type(path, ['Документ о качестве'])
    data_qualdoc = fun.sums_rows_by_keys(data_qualdoc,
                                         ['Марка', 'Наименование'],
                                         ['Количество', 'Вес общий'])

    # Получаем данные из КМД
    data_kmd = excel_KMD.get_data(path)

    # Получаем данные из АООК
    data_aook, date_aook = excel_AOOK.get_data(path,'*документ о качестве стальных*', 'Качества')

    # Сверяем данные
    result = rec__KMD_quardoc.reconciliation_data(data_kmd, data_qualdoc, data_aook, date_aook)

    # Вывод данныех в Excel
    out__KMD_qualdoc.data_output(result, FAIL_NAME)

    # Добовляем данные в книгу
    # xl(FAIL_NAME, 'Качество->КМД', data_qualdoc)
    # xl(FAIL_NAME, 'КМД->Качество', data_kmd)
    # xl(FAIL_NAME, 'АООК->Качество', data_aook)


def main():
    try:
        # Удаляем файл, если он етсь
        if os.path.exists(FAIL_NAME):
            os.remove(FAIL_NAME)

        # # Выполняем сверки
        reconciles_documents_kmd_aorpi()
        reconciles_documents_jsr_csv()
        reconciles_documents_aorpi_qualdoc()
        reconciles_documents_kmd_qualdoc()
    except Exception as ex:
        # Настройка конфигурации логирования
        logging.basicConfig(filename=path + '\\Отчет об ошибках.txt', level=logging.ERROR,
                            format='%(asctime)s - %(levelname)s - %(message)s')
        print('Ошибка:\n', traceback.format_exc())
        logging.exception("Произошла ошибка: %s", str(ex))
        return exit(1)


if __name__ == "__main__":
    main()
