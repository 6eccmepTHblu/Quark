from Class import Excel_styles as styles
from def_folder.excel_collection import get_list_in_excel, enter_data_on_the_sheet
from consts.colors import COLORS
from openpyxl.utils import get_column_letter as collumn


SHEET_NAME = "Отчет"
HEADERS = {'ФИО': 'ФИО',
           'Дата поступления': 'Дата поступления в ОК',
           'Дата в бухгалтерию': 'Дата поступления в бухгалтерию (дата проведения)',
           'Номер удостоверения': '№ удостоверения',
           'Дата регестрации': 'Дата регестрации удостоверения',
           'Место': 'Место назначения (область, город, поселок)',
           'Дата убытия': 'Дата убытия',
           'Дата прибытия': 'Дата прибытия из служебной поездки',
           'Сумма': 'Сумма аванса, руб.',
           'Примечание': 'Примечание',
           'ТаблицаИтого': 'Таблица отчета'}

def data_output(table, file_name):
    if not table:
        print('!!! Сврека Качества с АоРПИ отсутствует, лист не будет создан')
        return None

    wb, sh = get_list_in_excel(file_name, SHEET_NAME)

    enter_data_on_the_sheet(table, HEADERS, sh)

    # Настройка визуальной части
    prefix_style = 'report'

    # Заголовок
    header_style = styles.RangeStyle(sh, 'Header' + prefix_style, COLORS['Синий светлый'], size=13, bold=True)
    header_style.apply_style(c2=len(HEADERS)-1)

    # Таблица
    table_style = styles.RangeStyle(sh, 'Table' + prefix_style)
    table_style.apply_style(r1=2, r2=len(table), c2=len(HEADERS)-1)

    # Общие
    styles.add_filter(sh, row2=len(table), col2=len(HEADERS)-1)
    styles.SheetStyle(sh, 35, [1, 2], 85, COLORS['Синий светлый'])

    # Таблица итогов
    sh.column_dimensions[collumn(len(HEADERS))].hidden = True

    # Сохранить книгу
    try:
        wb.save(file_name)
        print('Книга сохранена - "' + file_name + '"')
    except:
        print('ФАЙЛ ЗАКРОЙ!!!')
