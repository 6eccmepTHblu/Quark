from Class import Excel_styles as styles

from def_folder.excel_collection import get_number_key_in_dict as get_key
from def_folder.excel_collection import get_list_in_excel, enter_data_on_the_sheet
from def_folder.excel_style import (get_column_by_key as col,
                                    conditional_formatting as uf_color)

from data_reconciliation.rec__KMD_AoRPI import STATUS
from consts.colors import COLORS

SHEET_NAME = "АоРПИ_КМД_АООК"
HEADERS = {'Марка': 'Марка',
           'Наименование': 'Наименование из КМД',
           'Количество': 'Кол-во КМД',
           'Статус проверки': 'Результат',
           'АоРПИ Номер из АООК': 'Упоминание в АООК',
           'АООК Дата': 'Дата АООК'}

def data_output(table, file_name):
    if not table:
        print('!!! Сврека КМД с АоРПИ отсутствует, лист не будет создан')
        return None

    wb, sh = get_list_in_excel(file_name, SHEET_NAME)

    enter_data_on_the_sheet(table, HEADERS, sh)

    # Настройка визуальной части
    prefix_style = 'kmd'

    # Заголовок
    header_style = styles.RangeStyle(sh, 'Header_' + prefix_style, COLORS['Синий светлый'], size=13, bold=True)
    header_style.apply_style(c2=len(HEADERS))

    # Таблица
    table_style = styles.RangeStyle(sh, 'Table_' + prefix_style, cell_color=COLORS['Зелёный светлый'])
    table_style.apply_style(r1=2, r2=len(table), c2=len(HEADERS))

    # Результат
    result_style = styles.RangeStyle(sh, 'Result_' + prefix_style, cell_color=COLORS['Красный'], bold=True)
    result_style.apply_style(r1=2, c1=get_key(HEADERS, 'Статус проверки'), r2=len(table),
                             c2=get_key(HEADERS, 'АоРПИ Номер из АООК'))

    # Общие
    styles.add_filter(sh, row2=len(table), col2=len(HEADERS))
    styles.SheetStyle(sh, 65, [1, 2], 100, COLORS['Синий светлый'])

    # Условное ворматрирование
    uf = {
        f'AND(ROW()>1, SEARCH("{STATUS["Полное совпадение"]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Зелёный'], 'Диапазон': col(HEADERS, ["Статус проверки", "Статус проверки"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["Наименование"]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Жёлтый'], 'Диапазон': col(HEADERS, ["Наименование", "Наименование"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["Количество"]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Жёлтый'], 'Диапазон': col(HEADERS, ["Количество", "Количество"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["Марка"]}",{col(HEADERS, "Статус проверки")}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["Марка", "Количество"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК != дата"]}",{col(HEADERS, "АоРПИ Номер из АООК")}))': [
            {'Цвет': COLORS['Жёлтый'], 'Диапазон': col(HEADERS, ["АоРПИ Номер из АООК", "АоРПИ Номер из АООК"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК ="]}",{col(HEADERS, "АоРПИ Номер из АООК")}))': [
            {'Цвет': COLORS['Зелёный'], 'Диапазон': col(HEADERS, ["АоРПИ Номер из АООК", "АоРПИ Номер из АООК"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК !="]}",{col(HEADERS, "АоРПИ Номер из АООК")}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["АоРПИ Номер из АООК",
                                                                               "АоРПИ Номер из АООК"])}
        ],
        f'AND(ROW()>1, SEARCH("{STATUS["АООК < АоРПИ"]}",{col(HEADERS, "АООК Дата")}))': [
            {'Цвет': COLORS['Красный светлый'], 'Диапазон': col(HEADERS, ["АООК Дата", "АООК Дата"])}
        ]
    }
    uf_color(uf, sh)

    # Сохранить книгу
    try:
        wb.save(file_name)
        print('Книга сохранена - "' + file_name + '"')
    except:
        print('ФАЙЛ ЗАКРОЙ!!!')