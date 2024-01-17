from Class import Excel_styles as styles

from def_folder.excel_collection import get_number_key_in_dict as get_key
from def_folder.excel_collection import get_list_in_excel, enter_data_on_the_sheet
from def_folder.excel_style import (get_column_by_key as col,
                                    conditional_formatting as uf_color)

from data_reconciliation.rec__DOC__DK_AoRPI import STATUS
from consts.colors import COLORS

SHEET_NAME = "Комплектность"
HEADERS = {'Номер документа': 'Номер ДК',
           'Номер док.': 'Номера ДК в АоРПИ',
           'Дата отгрузки': 'Дата\nАоРПИ',
           'Статус проверки': 'Результат'}


def data_output(table, file_name):
    if not table:
        print('!!! Сврека ДК с АоРПИ отсутствует, лист не будет создан')
        return None

    wb, sh = get_list_in_excel(file_name, SHEET_NAME, one_sheet=True)
    ind = sh.max_column + 2

    enter_data_on_the_sheet(table, HEADERS, sh, c1=ind)

    # Настройка визуальной части
    prefix_style = 'дк_аорпи'

    # Заголовок
    header_style = styles.RangeStyle(sh, 'Header_' + prefix_style, COLORS['Синий светлый'], size=13, bold=True)
    header_style.apply_style(c1=ind, c2=ind+len(HEADERS)-1)

    # Таблица
    table_style = styles.RangeStyle(sh, 'Table_' + prefix_style, cell_color=COLORS['Зелёный светлый'])
    table_style.apply_style(r1=2, r2=len(table), c1=ind, c2=ind+len(HEADERS)-1)

    # Результат
    result_style = styles.RangeStyle(sh, 'Result_' + prefix_style, cell_color=COLORS['Зелёный'], bold=True)
    result_style.apply_style(r1=2, c1=ind+get_key(HEADERS, 'Статус проверки')-1, r2=len(table),
                             c2=ind+get_key(HEADERS, 'Статус проверки')-1)

    # Общие
    styles.SheetStyle(sh, 65)

    # Условное ворматрирование
    uf = {
        f'AND(ROW()>1, SEARCH("{STATUS["АООК !="]}",{col(HEADERS, "Статус проверки", ind)}))': [
            {'Цвет': COLORS['Красный светлый'],
             'Диапазон': col(HEADERS, ["Номер док.", "Номер док."], ind)}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АОРПИ !="]}",{col(HEADERS, "Статус проверки", ind)}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["Номер документа", "Номер документа"], ind)}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК != дата"]}",{col(HEADERS, "Статус проверки", ind)}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["Дата отгрузки", "Дата отгрузки"], ind)}
        ],
        f'AND(ROW()>1, {col(HEADERS, "Статус проверки", ind)}<>"")': [
            {'Цвет': COLORS['Красный'], 'Диапазон': col(HEADERS, ["Статус проверки", "Статус проверки"], ind)}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК < АоРПИ"]}",{col(HEADERS, "Статус проверки", ind)}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["Дата отгрузки", "Дата отгрузки"], ind)}
        ]
    }
    uf_color(uf, sh)

    # Сохранить книгу
    try:
        wb.save(file_name)
        print('Книга сохранена - "' + file_name + '"')
    except:
        print('ФАЙЛ ЗАКРОЙ!!!')
