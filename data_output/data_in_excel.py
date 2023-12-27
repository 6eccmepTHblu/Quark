import os
import openpyxl

from Class import Excel_styles as styles

from def_folder.excel_collection import enter_data_on_the_sheet
from consts.colors import COLORS


def main(path, sheet_name, table):
    if not table:
        print(f'!!! Данные "{sheet_name}" отсутствуют, лист не будет создан!')
        return None

    # Проверяем, существует ли файл
    if not os.path.exists(path):
        print(f'Книга "{path}" не найдена')
        return ''

    wb = openpyxl.load_workbook(path)  # Подгрузить гингу
    sh = wb.create_sheet(sheet_name)  # Создание листа

    headers_key = set()
    for row in table:
        for kay in row.keys():
            headers_key.add(kay)
    headers_key = {key: key for key in headers_key}

    # Выводим данные
    enter_data_on_the_sheet(table, headers_key, sh, False)

    # Настройка визуальной части_______________________________________________
    # Заголовок
    header_style = styles.RangeStyle(sh, 'Header_' + sheet_name, COLORS['Синий светлый'], size=13, bold=True)
    header_style.apply_style(c2=len(table[0]))

    # Таблица
    table_style = styles.RangeStyle(sh, 'Table_' + sheet_name)
    table_style.apply_style(r1=2, r2=len(table), c2=len(table[0]))

    # Общие
    styles.add_filter(sh, row2=len(table)+1, col2=len(table[0]))
    styles.SheetStyle(sh, 65, [1, 2], 85, COLORS['Зелёный светлый'])

    sh.sheet_state = 'hidden'  # Скрыть лист

    try:
        wb.save(path)  # Сохраняем книгу
        print('  Лист "' + sheet_name + '" создан!')
    except:
        print('!!! Не удалось создать лист "' + sheet_name + '"')
