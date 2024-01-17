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

# Пример использования:
list_of_dicts = [
    {'a': 1, 'b': 2},
    {'b': 3, 'c': 4},
    {'a': 5, 'c': 6}
]

result = fill_missing_keys(list_of_dicts)
print(result)