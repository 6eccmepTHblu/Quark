FILES = {
    'ВИК Лимак': {
        'Маска': ['ВИК Лимак.csv'],
        'PDF': 'ВИК Лимак',
        'Заголовки': {
            'Номер': [0, 'Номер'],
            'Дата': [1, 'Дата'],
            'Файл': 11,
            'СтраницаОбщая': 15
        },
        'Таблица': 'ВИК Лимак таблица'
    },
    'ВИК Лимак таблица': {
        'Маска': ['table_data\\table_ВИК Лимак_Таблица_*.csv'],
        'Заголовки': {
            'НаимДетали': [1, 'Деталь'],
            'Толщена': [2, 'Толщена'],
            'Материал': [3, 'Материал'],
            'ТипСварки': [4, 'Тип cварки'],
            'ФИО/Клеймо': [5, 'Клеймо'],
            'Заключение': [7, 'Заключение']
        },
        'Отступ': 2
    },
    'ПВК Лимак': {
        'Маска': ['ПВК Лимак.csv'],
        'PDF': 'ПВК Лимак',
        'Заголовки': {
            'Номер': [0, 'Номер'],
            'Дата': [1, 'Дата'],
            'Файл': 11,
            'СтраницаОбщая': 15
        },
        'Таблица': 'ПВК Лимак таблица'
    },
    'ПВК Лимак таблица': {
        'Маска': ['table_data\\table_ПВК Лимак_Таблица_*.csv'],
        'Заголовки': {
            'НаимДетали': [1, 'Деталь'],
            'Толщена': [2, 'Толщена'],
            'Материал': [3, 'Материал'],
            'ТипСварки': [4, 'Тип cварки'],
            'ФИО/Клеймо': [5, 'Клеймо'],
            'Заключение': [7, 'Заключение']
        }
    },
    'АОРПИ Усть-Луга': {
        'Маска': ['АОРПИ Усть-Луга.csv'],
        'PDF': 'АОРПИ Усть-Луга',
        'Заголовки': {
            'Номер': [0, 'Номер', 'НомерАоРПИ'],
            'Дата': [4, 'Дата'],
            'Файл': 23,
            'СтраницаОбщая': 27
        },
        'Таблица': ['АОРПИ Усть-Луга изделия',
                    'АОРПИ Усть-Луга сопр.док.']
    },
    'АОРПИ Усть-Луга изделия': {
        'Маска': ['table_data\\table_АОРПИ Усть-Луга_Таблица изделие_*.csv'],
        'Заголовки': {
            'НаимПродукции': [0, 'НаимПродукции'],
            'Марка': [1, 'Марка'],
            'Количество': [4, 'Количество']
        },
    },
    'АОРПИ Усть-Луга сопр.док.': {
        'Маска': ['table_data\\table_АОРПИ Усть-Луга_Таблица сопроводительная документация_*.csv'],
        'Заголовки': {
            'Тип док.': [0, 'Тип док'],
            'Номер док.': [1, 'Номер'],
            'Дата док.': [2, 'Дата']
        }
    },
    'Документ о качестве': [
        {
            'Маска': ['Документ о качестве Металлострой.csv'],
            'PDF': 'Документ о качестве Металлострой',
             'Заголовки': {
                 'Номер документа': [0, 'Номер'],
                 'Дата отгрузки': [6, 'Дата'],
                 'Файл': 13,
                 'СтраницаОбщая': 17
             },
             'Таблица': 'Документ о качестве таблица Металлострой'
         },
        {
            'Маска': ['Документ о качестве DiFerr.csv'],
            'PDF': 'Документ о качестве DiFerr',
             'Заголовки': {
                 'Номер документа': [0, 'Номер'],
                 'Дата отгрузки': [1, 'Дата'],
                 'Файл': 14,
                 'СтраницаОбщая': 18
             },
             'Таблица': 'Документ о качестве таблица DiFerr'
         }
    ],
    'Документ о качестве таблица Металлострой': {
        'Маска': ['table_data\\table_Документ о качестве Металлострой_Таблица Конструкции_*.csv'],
        'Заголовки': {
            'Наименование': [2, 'НаимПродукции'],
            'Марка': [1, 'Марка'],
            'Количество': [3, 'Количество'],
            'Вес общий': [5, 'Вес']
        },
        'Отступ': 1
    },
    'Документ о качестве таблица DiFerr': {
        'Маска': ['table_data\\table_Документ о качестве DiFerr_Таблица Конструкции_*.csv'],
        'Заголовки': {
            'Наименование': [2, 'НаимПродукции'],
            'Марка': [1, 'Марка'],
            'Количество': [4, 'Количество'],
            'Вес общий': [5, 'Вес']
        },
        'Отступ': 1
    },
    'Авансовый отчет': {
        'Маска': ['Авансовый отчет.csv'],
        'PDF': 'Авансовый отчет',
        'Заголовки': {
            'ФИО': 0,
            'Дата убытия': [1, 'Дата'],
            'Дата 1': [2, 'Дата'],
            'Дата 2': [3, 'Дата'],
            'Дата 3': [4, 'Дата'],
            'Место': 5,
            'Файл': 7,
            'СтраницаОбщая': 11
        },
        'Таблица': 'Авансовый отчет таблица'
    },
    'Авансовый отчет таблица': {
        'Маска': ['table_data\\table_Авансовый отчет_Таблица_*.csv'],
        'Заголовки': {
            'Номер': 0,
            'Дата': [1, 'Дата'],
            'Описание': 2,
            'Сумма': [3, 'Сумма']
        },
        'Отступ': 1
    }
}