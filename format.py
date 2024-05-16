import pandas as pd

file = pd.read_excel('AI4Merge.xlsx', sheet_name=None, engine='openpyxl')
for key in file.keys():
    file[key].rename(columns={'교육생_x': '교육생'}, inplace=True)
    print(file[key])