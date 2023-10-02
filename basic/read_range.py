import openpyxl

# Buka file Workbook
workbook = openpyxl.load_workbook("sample.xlsx")

# Pilih Worksheet "gaji"
worksheet = workbook["GAJI"]

# Inisialisasi variabel untuk menyimpan angka desimal yang ditemukan
angka_desimal = []

# Loop melalui sel A1 hingga C15
for row in worksheet.iter_rows(min_row=2, max_row=46, min_col=1, max_col=5):
    for cell in row:
        if isinstance(cell.value, (float, int)):
            # Cek apakah nilai sel adalah angka desimal atau integer
            if cell.value % 1 != 0:
                angka_desimal.append(cell.value)

# Cetak angka desimal yang ditemukan
print("Angka desimal yang ditemukan:")
x = 0
for angka in angka_desimal:
    x = x + 1
    print(f"{x} {angka}")

# Tutup file Workbook
workbook.close()
