import mysql.connector
def display():

    # Konfigurasi koneksi ke database
    config = {
        'user': 'rangga',
        'password': 'rangga',
        'host': 'localhost',
        'database': 'sekolah',
        'raise_on_warnings': True
    }

    try:
        # Membuat koneksi ke database
        connection = mysql.connector.connect(**config)

        if connection.is_connected():
            print("Berhasil terhubung ke database")

            # Membuat cursor untuk eksekusi query
            cursor = connection.cursor()

            # Query untuk mengambil data dari tabel employees
            query = "SELECT Nis, Nama, Umur, Seks FROM siswa"
            cursor.execute(query)

            # Mengambil semua baris data hasil eksekusi query
            rows = cursor.fetchall()

            print(f"No  Nis  Nama  Umur  Gender")

            # Menampilkan data
            i = 0
            for row in rows:
                i += 1
                print(f"{i}  {row[0]}  {row[1]}  {row[2]}  {row[3]}")

    except mysql.connector.Error as err:
        print(f"Error: {err}")

    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()
            print("Koneksi ke database ditutup")
