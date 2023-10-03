import pandas as pd
import numpy as np

# Baca file Excel
path = input('Negative Values in Worksheet\nFile Name: ')

worksheet = pd.ExcelFile(path)
sheetName = worksheet.sheet_names
print("Available Worksheet(s): ")
for name in sheetName:
    print(f'  {name}')

wsheet = input('Get 1 Worksheet: ')

df = pd.read_excel(path, sheet_name=wsheet, usecols="A:E")

# Inisialisasi variabel untuk menyimpan hasil pencarian
hasil_pencarian = []

# Loop melalui baris dan kolom dalam DataFrame
for index, row in df.iterrows():
    for column, value in row.items():
        if isinstance(value, float):
            if value % 1 != 0 and not np.isnan(value):
                # Cek apakah nilai sel adalah angka desimal
                alamat = f"{df.columns.get_loc(column) + 1}{index + 1}"
                hasil_pencarian.append((alamat, value))

# Cetak hasil pencarian
print("Hasil pencarian angka Decimal/Float/Double:\n\n No Clm Row Value")
x = 0
for alamat, nilai in hasil_pencarian:
    x = x + 1
    kolom = alamat[0]
    baris = alamat[1]
    print(f" {x}   {kolom}  {baris}: {nilai}")
