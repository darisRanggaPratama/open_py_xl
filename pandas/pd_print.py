import pandas as pd

# Baca file Excel
path = input('Display Values in Worksheet\nFile Name: ')

worksheet = pd.ExcelFile(path)
sheetName = worksheet.sheet_names
print("Available Worksheet(s): ")
for name in sheetName:
    print(f'  {name}')

wsheet = input('Get 1 Worksheet: ')

df = pd.read_excel(path, sheet_name=wsheet, usecols='A:F')

print(df)