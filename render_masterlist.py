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

def compute_layout(
    width=1080,
    height=1920,
    margin=0,
    gutter=0,
    header_h=70,
    pairs_per_row=3,
    rows_per_page=18,
    name_ratio=0.72
):
    # calculate usable area to contain all the rows (including header)
    usable_w = width - 2 * margin
    usable_h = height - 2 * margin
    
    # width split for each pair
    pair_w = (usable_w - (pairs_per_row - 1) * gutter) // pairs_per_row
    # use `int` here just to make sure it doesnt give decimal number
    name_w = int(pair_w * name_ratio)
    brn_w = pair_w - name_w
    
    # height split (uniform rows, header absorbs the remaining px)
    body_h_raw = usable_h - header_h 
    row_h = body_h_raw // rows_per_page
    body_h = row_h * rows_per_page
    leftover = body_h_raw - body_h
    header_h = header_h + leftover
    
    row_height = [row_h] * rows_per_page
    
    row_tops = []
    y = margin + header_h
    
    for h in row_height:
        row_tops.append(y)
        y += h
        
    return {
        'usable_w' : usable_w,
        'usable_h' : usable_h,
        'pair_w' : pair_w,
        'name_w' : name_w,
        'brn_w' : brn_w,
        'header_h' : header_h,
        'row_h' : row_h,
        'row_height_sum' : sum(row_height),
        'row_top3_position' : row_tops[:3],
        'row_last3_position' : row_tops[-3:],
    }

# path = 'masterlist.xlsx'

# all_rows = []

# sheets = get_sheet_names(path)
# for sheet in sheets:
#     rows = load_rows_from_sheet(path, sheet)
    # we use extend instead of append to make sure it keeps a flat list instead of nested list
    # all_rows.extend(rows)

# print('Total rows: ', len(all_rows))
# `list[start:stop]`
# print('First 5 rows: ', all_rows[:5])
# print('Last 5 rows: ', all_rows[-5:])

layout = compute_layout()

for k,v in layout.items():
    print(k, ':', v)