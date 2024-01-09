import pandas as pd

col_1 = {'name1': 'col_name', 'date1': 'col_date'}
col_2 = {'name2': 'col_name', 'date2': 'col_date'}
key_1 = 'name1'
key_2 = 'name2'

table_1 = [
    {'col_name': 'номер1', 'col_date_1': '01.01.21', 'col_date_2': '01.01.21'},
    {'col_name': 'номер2', 'col_date_1': '01.01.21', 'col_date_2': '02.01.21'},
    {'col_name': 'номер3', 'col_date_1': '01.01.21', 'col_date_2': '01.01.21'},
    {'col_name': 'номер4', 'col_date_1': '03.01.21', 'col_date_2': '03.01.21'},
]

table_2 = [
    {'col_name': 'номер1', 'col_date': '01.01.21'},
    {'col_name': 'номер3', 'col_date': '03.01.21'},
    {'col_name': 'номер4', 'col_date': '04.01.21'},
    {'col_name': 'номер5', 'col_date': '05.01.21'},
]

# Преобразование таблиц в DataFrame
df_1 = pd.DataFrame(table_1)
df_2 = pd.DataFrame(table_2)

# Объединение по ключу "col_name"
merged_df = pd.merge(df_1, df_2, left_on=col_1[key_1], right_on=col_2[key_2], how='outer', suffixes=('_1', '_2'))

# Вывод результата
print(merged_df)