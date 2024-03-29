import logging

from Class import Excel_styles as styles

from def_folder.data_collection import get_number_key_in_dict as get_key
from def_folder.data_collection import get_list_in_excel, enter_data_on_the_sheet
from def_folder.excel_style import (get_column_by_key as col,
                                    conditional_formatting as uf_color)

from data_reconciliation.rec__DOC__AoRPI_AOOK import STATUS
from consts.colors import COLORS

SHEET_NAME = "Комплектность"
HEADERS = {'Номер': 'Номер АОРПИ',
           'АоРПИ Номер из АООК': 'Номера АоРПИ в АООК/АОСР',
           'Дата': 'Дата АОРПИ',
           'Статус проверки': 'Результат'}

def data_output(table, file_name, date_aook):
    if not table:
        logging.warning(f"Сврека АООК с АоРПИ отсутствует, лист не будет создан!")
        return None

    wb, sh = get_list_in_excel(file_name, SHEET_NAME, one_sheet=True)
    ind = sh.max_column + 2

    HEADERS['Дата'] += str(f'\nАООК/АОСР {date_aook}')
    enter_data_on_the_sheet(table, HEADERS, sh)

    # Настройка визуальной части
    prefix_style = 'аорпи_аоок'

    # Заголовок
    header_style = styles.RangeStyle(sh, 'Header_' + prefix_style, COLORS['Синий светлый'], size=13, bold=True)
    header_style.apply_style(c2=len(HEADERS))

    # Таблица
    table_style = styles.RangeStyle(sh, 'Table_' + prefix_style, cell_color=COLORS['Зелёный светлый'])
    table_style.apply_style(r1=2, r2=len(table), c2=len(HEADERS))

    # Результат
    result_style = styles.RangeStyle(sh, 'Result_' + prefix_style, cell_color=COLORS['Красный'], bold=True)
    result_style.apply_style(r1=2, c1=get_key(HEADERS, 'Статус проверки'), r2=len(table),
                             c2=get_key(HEADERS, 'Статус проверки'))

    # Общие
    styles.add_filter(sh, row2=len(table), col2=len(HEADERS))
    styles.SheetStyle(sh, 65, [1, 2], 100, COLORS['Синий светлый'])

    # Условное ворматрирование
    uf = {
        f'AND(ROW()>1, SEARCH("{STATUS["АОРПИ !="]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["АоРПИ Номер из АООК", "АоРПИ Номер из АООК"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК !="]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["Номер", "Номер"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК != дата"]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["Дата", "Дата"])}
        ],
        f'AND(ROW()>1, {col(HEADERS, "Статус проверки")}="{STATUS["Совпало"]}")': [
            {'Цвет': COLORS['Зелёный'], 'Диапазон': col(HEADERS, ["Статус проверки", "Статус проверки"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК < АоРПИ"]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["Дата", "Дата"])}
        ]
    }
    uf_color(uf, sh)

    # Сохранить книгу
    try:
        wb.save(file_name)
        print('Книга сохранена - "' + file_name + '"')
    except:
        print('ФАЙЛ ЗАКРОЙ!!!')
