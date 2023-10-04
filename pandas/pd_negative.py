import numpy as np
import pandas as pd

# Baca file Excel
path = input('Negative Values in Worksheet\nFile Name: ')

worksheet = pd.ExcelFile('sample.xlsx')
sheetName = worksheet.sheet_names
print("Available Worksheet(s): ")
for name in sheetName:
    print(f'  {name}')

wsheet = input('Get 1 Worksheet: ')

df = pd.read_excel(path, sheet_name=wsheet, usecols='A:E')

# Inisialisasi variabel untuk menyimpan hasil pencarian
hasil_pencarian = []

# Loop melalui baris dan kolom dalam DataFrame
for index, row in df.iterrows():
    for column, value in row.items():
        if isinstance(value, float):
            # Cek apakah nilai sel adalah angka desimal
            if value < 0 and not np.isnan(value):
                hasil_pencarian.append((df.columns.get_loc(column), index, value))

# Cetak hasil pencarian
print('Hasil pencarian angka Negative:\n\n No Row Clm Value')
x = 0
for kolom, baris, nilai in hasil_pencarian:
    x = x + 1
    print(f' {x}  {hash(baris) + 2}   {kolom + 1}:   {nilai}')
