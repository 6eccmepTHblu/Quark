import openpyxl
import fnmatch

from def_folder import data_collection as fun
from def_folder import data_normalization as norm
from def_folder import excel_collection as excel

DATE_CELL = '*(дата составления акта)*'
SHEET_REESTR = '*реестр*'
TEMPLATES = ['*АООК*.xls*',
             '*Акт освидетельствования ответственных конструкций*.xls*']
HEADERS = {'Пункт': '№ п/п',
           'ИД': '*документ*',
           'Номер/Дата': '*Реквизит*'}


def get_data(path, mask, type_name):
    # Найти файл
    file = fun.find_file(path, TEMPLATES)
    if file == '':
        return None

    # Открываем книгу АООК
    wb = openpyxl.load_workbook(file)

    # Находим дату АООК
    sh = wb[excel.find_sheet(wb, '*АООК*')[0]]
    data_aook = ''
    for row in sh.iter_rows():
        for cell in row:
            if fnmatch.fnmatch(str(cell.value), DATE_CELL):
                # Если нашли ячейку со значением "дата", то берем значение из строки выше
                data_aook = sh.cell(row=cell.row - 1, column=cell.column).value
                break
        if data_aook:
            break

    # Нормируем дату АООК
    data_aook = data_aook.replace('"', '').replace('г.', '').strip()
    data_aook = norm.get_date(data_aook)
    print("Дата АООК:", data_aook)

    # Собираем реестры АООК
    reestr_aook = []
    for sh in wb:
        if fnmatch.fnmatch(sh.title.lower(), SHEET_REESTR):
            all_data = excel.get_data_excel(sh)

            # Ищем нужные столбцы
            headers, resul_rows = excel.find_headers(all_data, HEADERS)

            # Выбераем нужные столбцы, разбивая по словарю
            data = []
            for row in all_data[resul_rows:]:
                row_data = {}
                for key, item in headers.items():
                    row_data[key] = row[item]
                data.append(row_data)

            # Фильтруем данные по АоРПИ
            for row in data:
                if fnmatch.fnmatch(str(row['ИД']).lower(), mask):
                    reestr_aook.append(row)

    # Нормирование
    for row in reestr_aook:
        if 'Номер/Дата' in row:
            # Дата
            date = row['Номер/Дата'].split('от')
            if len(date) > 1:
                date = date[1]
            else:
                date = date[0]
            row[f'{type_name} Дата из АООК'] = norm.get_date(date)
            # Номер
            nambe = row['Номер/Дата'].split('от')[0].strip().replace('№', '')
            row[f'{type_name} Номер из АООК'] = norm.transliteration_ru_en(nambe).upper()

    print('Данные из АООК собранны - ' + str(len(reestr_aook)) + '.')
    return reestr_aook, data_aook


if __name__ == '__main__':
    get_data('D:\Работа\Силенко Д.Т\Задача 2(Путхон - Кварк V2)\Данные')
