import pandas as pd

# Загружаем данные из файла
base_df = pd.read_excel('base.xlsx')
converted_base_df = pd.read_excel('converted_base.xlsx')

for i in converted_base_df.index:
    #print(f"i '{i}' ' {base_df.iloc[i, 1]}")
    print(i)

