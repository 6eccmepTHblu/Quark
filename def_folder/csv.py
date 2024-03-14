import os
import csv


def get_data_in_csv(list_csv: list, name: bool = True, take_strio: bool = True) -> list:
    """
    **Функция `get_data_in_csv`**
    Эта функция считывает данные из файлов CSV, объединяет их в список и возвращает полученные данные.

    **Параметры**
    - `list_csv` (список): Список имен файлов CSV для чтения данных.
    - `name` (булево значение): Если установлено в `True`, добавляет в начало списка имя файла.
            По умолчанию установлено в `True`.

    **Возвращаемое значение**
    - `element_list` (список): Список списков, представляющих данные из файлов CSV.
            Если `name` установлено в `True`, каждый подсписок начинается с имени файла.

    **Пример использования**
    csv_files = ["file1.csv", "file2.csv"]
    data_list = get_data_in_csv(csv_files, indent=2, name=True)
    """
    element_list = []
    if not list_csv:
        return element_list

    if isinstance(list_csv, str):
        list_csv = [list_csv]

    for file_name in list_csv:
        if not os.path.exists(file_name):
            print(f"Файл не существует: {file_name}")
            continue

        with open(file_name) as my_file:
            try:
                file_contents = csv.reader(my_file, delimiter=';')
                if take_strio:
                    temp_list = [[str(cell).strip() for cell in row] for row in file_contents]
                else:
                    temp_list = [[str(cell) for cell in row] for row in file_contents]
                if name:  # Если name = True, добавляет в начало списка имя файла
                    temp_list = [row + [file_name] for row in temp_list]
                element_list.extend(temp_list[:])
            except csv.Error as e:
                print(f"Ошибка при чтении файла {file_name}: {e}")

    return element_list