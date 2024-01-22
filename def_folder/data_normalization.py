import fnmatch
import os
import re


def take_out_the_stamp(text):
    stamp = text
    if text:
        re_temp = r'[A-Z0-9]{4}'
        stamp = ''
        if re.search(re_temp, str(text)):
            stamp = re.search(re_temp, str(text)).group(0)
    return stamp


def transliteration_ru_en(text, type='en'):
    result = text
    if text:
        dictionary = {
            'а': 'a', 'с': 'c', 'е': 'e', 'к': 'k', 'м': 'm', 'о': 'o',
            'р': 'p', 'и': 'u', 'у': 'y', 'х': 'x',
            'А': 'A', 'В': 'B', 'Е': 'E', 'К': 'K', 'М': 'M', 'О': 'O',
            'Р': 'P', 'С': 'C', 'Т': 'T', 'Н': 'H', 'У': 'Y', 'Х': 'X'
        }
        if type == 'ru':
            dictionary = {v: k for k, v in dictionary.items()}
        result = ''.join([dictionary.get(char, char) for char in str(text)])
    return result


def filter_data(table, header, masks=''):
    date = []
    if table and header:
        for row in table:
            find = True
            for mask in masks:
                if fnmatch.fnmatch(row.get(header, ''), mask):
                    find = False
                    break
            if find:
                date.append(row)
    return date


def append_value(text, value, symbol='\n'):
    if not text:
        return value
    if not value:
        return text
    return text + symbol + value


def delete_symbol_in_start(text, symbol):
    if text and symbol:
        if str(text).startswith(symbol):
            return text[1:]
    return text


def get_date(date_temp):
    date = normalize_date(date_temp)
    date = re.search(r"(\d+\.?\d+\.?\d+)", date)
    if date:
        return date.group(1)
    else:
        return '"' + date_temp + '"'


def normalize_date(date_temp):
    result = date_temp
    month_dict = {'января': '01', 'февраля': '02', 'марта': '03', 'апреля': '04', 'мая': '05', 'июня': '06',
                  'июля': '07', 'августа': '08', 'сентября': '09', 'октября': '10', 'ноября': '11', 'декабря': '12'}
    if date_temp:
        for key, item in month_dict.items():
            result:str = result.lower().replace(key,'.' + item + '.')
        result = result.replace('..','.').replace(' ','')
        result = remove_side_values(result, '[0-9]')

        date = re.search(r'[0-9]{8}', result)
        if date:
            result = date.group(0)[:2] + '.' + date.group(0)[2:4] + '.' + date.group(0)[4:]

        year = re.search(r'\.\d\d$', result)
        if year:
            result = result[:-len(year.group(0))] + '.20' + year.group(0).replace('.', '')

        date = re.search(r"(\d+\.?\d+\.?\d+)", str(result))
        if date:
            date = date.group(1).replace('.', '')
            result = date[:2] + '.' + date[2:4] + '.' + date[4:]

    return result


def splits_by_character(table, header, symbol):
    if table and header and symbol:
        data = []
        for row in table:
            divided = row[header].split(symbol)
            if len(divided) == 1:
                data.append(row)
            else:
                for elem in divided:
                    temp_dict = {}
                    for key, value in row.items():
                        temp_dict[key] = value
                    temp_dict[header] = elem
                    data.append(temp_dict)
        return data
    return table

def remove_side_values(text: str, regx: str = r'[а-яёА-ЯЁa-zA-Z0-9]') -> str:
    """
    **Функция `remove_side_values`**
    Эта функция удаляет символы с обеих сторон строки `text`,
    которые не соответствуют заданному регулярному выражению `regx`.

    **Параметры**
    - `text` (строка): Исходная строка, из которой нужно удалить символы с обеих сторон.
    - `regx` (строка): Регулярное выражение, которому должны соответствовать символы, чтобы быть сохраненными.
    По умолчанию установлено так, чтобы сохранять только буквы и цифры.

    **Возвращаемое значение**
    - `text` (строка): Строка, из которой удалены символы с обеих сторон,
        не соответствующие заданному регулярному выражению.

    **Пример использования**
    original_text = "123abc!@#$%^&*abc123"
    result_text = remove_side_values(original_text, r'[a-zA-Z0-9]')
    """
    while not re.search(r'^' + regx, text) and text != '':
        text = text[1:]
    while not re.search(regx + r'$', text) and text != '':
        text = text[:-1]
    return text

def fill_missing_keys(list_of_dicts):
    # Находим все уникальные ключи
    all_keys = set()
    for d in list_of_dicts:
        all_keys.update(d.keys())

    # Проходим заново по списку словарей и добавляем недостающие ключи с пустыми значениями
    for d in list_of_dicts:
        missing_keys = all_keys - set(d.keys())
        for key in missing_keys:
            d[key] = ''

    return list_of_dicts

def csv_to_pdf_path(csv_path, name):
    # Получение имени папки из пути к CSV-файлу
    folder_name = os.path.splitext(os.path.basename(csv_path))[0]

    # Формирование пути к PDF-файлу
    pdf_path = os.path.join(os.path.dirname(csv_path), folder_name, name)

    return pdf_path

def create_hyperlink(path_file, filename, name_pdf=None):
    if name_pdf is None:
        name_pdf = filename

    if filename == '':
        return f'=HYPERLINK("{path_file}", "{name_pdf}")'

    return f'=HYPERLINK("{csv_to_pdf_path(path_file, filename)}", "{name_pdf}")'