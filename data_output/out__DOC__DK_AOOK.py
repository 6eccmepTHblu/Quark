from Class import Excel_styles as styles

from def_folder.excel_collection import get_number_key_in_dict as get_key
from def_folder.excel_collection import get_list_in_excel, enter_data_on_the_sheet
from def_folder.excel_style import (get_column_by_key as col,
                                    conditional_formatting as uf_color)

from data_reconciliation.rec__DOC__AoRPI_AOOK import STATUS
from consts.colors import COLORS

SHEET_NAME = "Сверка документов 2"
HEADERS = {'Номер документа': 'Номер ДК',
           'Качества Номер из АООК': 'Номера ДК в АООК',
           'Дата отгрузки': 'Дата\nАОРПИ/АООК',
           'Статус проверки': 'Результат'}


def data_output(table, file_name):
    if not table:
        print('!!! Сврека АООК с АоРПИ отсутствует, лист не будет создан')
        return None

    wb, sh = get_list_in_excel(file_name, SHEET_NAME)

    enter_data_on_the_sheet(table, HEADERS, sh)

    # Настройка визуальной части
    prefix_style = 'дк_аоок'

    # Заголовок
    header_style = styles.RangeStyle(sh, 'Header_' + prefix_style, COLORS['Синий светлый'], size=13, bold=True)
    header_style.apply_style(c2=len(HEADERS))

    # Таблица
    table_style = styles.RangeStyle(sh, 'Table_' + prefix_style, cell_color=COLORS['Зелёный светлый'])
    table_style.apply_style(r1=2, r2=len(table), c2=len(HEADERS))

    # Результат
    result_style = styles.RangeStyle(sh, 'Result_' + prefix_style, cell_color=COLORS['Зелёный'], bold=True)
    result_style.apply_style(r1=2, c1=get_key(HEADERS, 'Статус проверки'), r2=len(table),
                             c2=get_key(HEADERS, 'Статус проверки'))

    # Общие
    styles.add_filter(sh, row2=len(table), col2=len(HEADERS))
    styles.SheetStyle(sh, 65, [1, 2], 100, COLORS['Синий светлый'])

    # Условное ворматрирование
    uf = {
        f'AND(ROW()>1, SEARCH("{STATUS["АОРПИ !="]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Красный светлый'],
             'Диапазон': col(HEADERS, ["Качества Номер из АООК", "Качества Номер из АООК"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК !="]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["Номер документа", "Номер документа"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК != дата"]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["Дата отгрузки", "Дата отгрузки"])}
        ],
        f'AND(ROW()>1, {col(HEADERS, "Статус проверки")}<>"")': [
            {'Цвет': COLORS['Красный'], 'Диапазон': col(HEADERS, ["Статус проверки", "Статус проверки"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК < АоРПИ"]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["Дата отгрузки", "Дата отгрузки"])}
        ]
    }
    uf_color(uf, sh)

    # Сохранить книгу
    try:
        wb.save(file_name)
        print('Книга сохранена - "' + file_name + '"')
    except:
        print('ФАЙЛ ЗАКРОЙ!!!')
