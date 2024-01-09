class DocumentsReconciliation():
    def __init__(self, table_1:list[dict], table_2:list[dict]):
        self.table_1 = table_1
        self.table_2 = table_2

        self.col_1 = None
        self.col_2 = None
        self.key_1 = None
        self.key_2 = None
        self.table_itog = []

    def __creates_table_itog(self):
        for row in self.table_1:
            row_itog = {}
            for key, item in self.col_1.items():
                row_itog[key] = row[item]
            self.table_itog.append(row_itog)

    def __crosses_table(self):
        self.__creates_table_itog()
        data_itog_dict = w[self.key_1]: row for row in self.table_itog}
        for row in self.table_2:
            if row.get(self.key_2, '+') in data_itog_dict:
                temp = data_itog_dict[row['АоРПИ Номер из АООК']]
                temp['Номер АООК'] = row['АоРПИ Номер из АООК']
                temp['Дата АООК'] = row['АоРПИ Дата из АООК']
            else:
                temp = {}
                temp['Номер АООК'] = row['АоРПИ Номер из АООК']
                temp['Дата АООК'] = row['АоРПИ Дата из АООК']
                data_itog.append(temp)

    def crosses_table(self, col_1:dict[str: str], col_2:dict[str: str], key_1:str, key_2:str):
        if len(col_1) != len(col_2):
            raise ValueError(f"Разная длина списков. \n 1 лист - {len(col_1)}, 2 лист {len(col_2)}")
        self.col_1 = col_1
        self.col_2 = col_2
        self.key_1 = key_1
        self.key_2 = key_2
        self.__crosses_table()
        return self.table_itog
