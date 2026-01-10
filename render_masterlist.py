import pandas as pd

path = 'masterlist.xlsx'
sheet = 'A'

df = pd.read_excel(path, sheet_name=sheet)

print(df.columns.tolist())
print(df.head())