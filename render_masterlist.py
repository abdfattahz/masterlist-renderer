import pandas as pd

path = 'masterlist.xlsx'
xl = pd.ExcelFile(path)
print(xl.sheet_names)