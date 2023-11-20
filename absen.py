!pip install openpyxl==3.1.2

# Import pandas and datetime modules
import pandas as pd
import datetime

# Import ast module for safe evaluation of user input
import ast

# Print the program title
print('=' * 30)
print('Program Absensi Sederhana')
print('=' * 30)

# Get the number of classes from the user input
jumlah_kelas = input('Masukkan Jumlah Kelas : ')

# Validate and convert the user input to an integer
try:
    jumlah_kelas = int(jumlah_kelas)
except ValueError:
    print('Input harus berupa angka')
    exit()

# Initialize an empty list to store the data
data = []

# Loop through each class
for x in range(1, jumlah_kelas + 1):
    # Print the class name
    print(f'X-{x}')
    # Get the number of attendance, absence, excused, and dispensation from the user input
    h = input('Masukkan Jumlah Kehadiran : ')
    i = input('Masukkan Jumlah Izin : ')
    s = input('Masukkan Jumlah Izin Sakit : ')
    d = input('Masukkan Jumlah Izin Dispensasi : ')

    # Validate and convert the user input to integers using ast.literal_eval
    try:
        h = ast.literal_eval(h)
        i = ast.literal_eval(i)
        s = ast.literal_eval(s)
        d = ast.literal_eval(d)
        if not (isinstance(h, int) and isinstance(i, int) and isinstance(s, int) and isinstance(d, int)):
            raise ValueError
    except (ValueError, SyntaxError):
        print('Input harus berupa angka')
        exit()

    # Append the data to the list as a dictionary
    data.append({'Kelas': f'X-{x}', 'Kehadiran': h, 'Izin': i, 'Izin Sakit': s, 'Izin Dispensasi': d})

# Create a pandas dataframe from the list of dictionaries
df = pd.DataFrame(data)

# Print the dataframe
print(df)

# Get the current date
today = datetime.date.today ()

# Format the date into dd-mm-yyyy
filename = today.strftime ("%d-%m-%Y")

# Add the .xlsx extension
filename = filename + ".xlsx"

# Export the dataframe to an excel file using a try-except block
try:
    df.to_excel(filename, index=False)
    # Print a success message
    print(f'Data berhasil diexport ke file {filename}')
except Exception as e:
    # Print an error message with the exception details
    print(f'Gagal mengekspor data ke file {filename}')
    print(f'Error: {e}')

# Raise an exception if the user input is not an integer
if not isinstance(jumlah_kelas, int):
    raise TypeError('Jumlah kelas harus berupa angka')
