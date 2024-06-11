import mysql.connector
import pandas as pd


def makeExcel():

    # Konfigurasi koneksi ke database
    config = {
        'user': 'rangga',
        'password': 'rangga',
        'host': 'localhost',
        'database': 'sekolah',
        'raise_on_warnings': True
    }

    try:
        connect = mysql.connector.connect(**config)
        if connect.is_connected():
            print("Connected to MySQL")

            # Crete cursor to execute query
            cursor = connect.cursor()
            sql = "SELECT Nis, Nama, Umur, Seks FROM siswa"
            cursor.execute(sql)
            # Get record
            rows = cursor.fetchall()
            # Close cursor
            cursor.close()

            # Convert query result to dataframe
            df = pd.DataFrame(rows, columns=['Nis', 'Nama', 'Umur', 'Seks'])

            # Save dataframe to excel
            file = "siswa.xlsx"
            df.to_excel(file, index=False)
            print(f"Query result is saved to Excel file: {file}")
    except mysql.connector.Error as err:
        print(f"Error while connecting to MySQL: {err}")
    finally:
        if connect.is_connected():
            connect.close()
            print("connection to database is closed")

