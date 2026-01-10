import pandas as pd

path = 'masterlist.xlsx'
xl = pd.ExcelFile(path)

for sheet in xl.sheet_names:
    # `nrows=0` so that it will only read the header
    df = pd.read_excel(path, sheet_name=sheet, nrows=0)
    print(sheet, '->', df.columns.tolist())