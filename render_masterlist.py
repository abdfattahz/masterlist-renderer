import pandas as pd

def get_sheet_names(path):
    xl = pd.ExcelFile(path)
    return xl.sheet_names

def load_rows_from_sheet(path, sheet_name):
    df = pd.read_excel(path, sheet_name=sheet_name)
    
    # advanded compact way
    # df.columns = [c.strip() for c in df.columns]

    # long version
    new_columns = []
    for c in df.columns:
        new_columns.append(c.strip())
    
    df.columns = new_columns
    
    df = df[['COMPANY NAME', 'COMPANY NO.']].dropna(how='all')

    df['COMPANY NAME'] = df['COMPANY NAME'].astype(str).str.strip()
    df['COMPANY NO.'] = df['COMPANY NO.'].astype(str).str.strip()
    
    # we do this so that each elaned rows will be output as tuples `('COMPANY NAME', 'COMPANY NO.')`
    return list(df.itertuples(index=False, name=None))