import openpyxl

# Buka file Workbook
workbook = openpyxl.load_workbook("sample.xlsx")

# Pilih Worksheet "gaji"
worksheet = workbook["GAJI"]

# Inisialisasi variabel untuk menyimpan alamat sel dan nilainya
hasil_pencarian = []

# Loop melalui sel A1 hingga C15
for row in worksheet.iter_rows(min_row=2, max_row=46, min_col=1, max_col=5):
    for cell in row:
        if isinstance(cell.value, float):
            # Cek apakah nilai sel adalah angka desimal
            hasil_pencarian.append((cell.coordinate, cell.value))

# print(type(hasil_pencarian))
# print(hasil_pencarian[1])
# Cetak hasil pencarian
print("Hasil pencarian angka Decimal/Float/Double:")
x = 0
for alamat, nilai in hasil_pencarian:
    x = x + 1
    print(f"{x} {alamat}: {nilai}")

# Tutup file Workbook
workbook.close()
