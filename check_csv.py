import pandas as pd

# Чтение CSV
df = pd.read_csv("D:/!tilo/COMCENTER/parsing_COMCENTER.RU/COMCENTER.ru_database/DATABASE_comcenter_products.csv")

# Проверка на NaN
print("Столбцы с NaN:")
print(df.isna().sum())

# Вывод строк с NaN
print("\nСтроки с NaN:")
print(df[df.isna().any(axis=1)])