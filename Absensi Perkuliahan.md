from tkinter import *
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

# Function to handle data insertion
def insert_data():
    global num
    num += 1
    sheetnum = num + 3

    data = [
        ("A" + str(sheetnum), num),
        ("B" + str(sheetnum), nama_entry.get()),
        ("C" + str(sheetnum), NIM_entry.get()),
        ("D" + str(sheetnum), jurusan_entry.get())
    ]

    for cell, value in data:
        sheet[cell] = value
        sheet[cell].font = font_bold
        sheet[cell].border = border
        sheet[cell].alignment = alignment_center

    sheet['B1'] = matkul_entry.get()
    sheet['B2'] = tanggal_entry.get()

    nama_entry.delete(0, END)
    NIM_entry.delete(0, END)
    jurusan_entry.delete(0, END)

# Function to save data to Excel
def save_data():
    filename = f"{matkul_entry.get()}_{tanggal_entry.get()}.xlsx"
    workbook.save(filename=filename)
    informasi.config(text=f"Data absen telah disimpan!\nNama file: {filename}")

# Function to reset data and information
def create_new_data():
    global num
    informasi.config(text='Klik Insert untuk semua mahasiswa, kemudian klik Save jika semua telah diabsen.')
    nama_entry.delete(0, END)
    NIM_entry.delete(0, END)
    jurusan_entry.delete(0, END)
    num = 0

# Function to calculate attendance percentage
def calculate_percentage():
    global workbook, sheet, font_bold, alignment_center, num
    try:
        filename = f"{matkul_entry.get()}_{tanggal_entry.get()}.xlsx"
        workbook = load_workbook(filename)
        sheet = workbook.active
        total_students = 7
        present_students = total_students - num

        percentage = (present_students / total_students) * 100
        informasi.config(text=f"Persentase kehadiran: {percentage:.2f}%")
        workbook.close()
    except Exception as e:
        informasi.config(text="Gagal menghitung persentase kehadiran.")

# Initialize the Tkinter window
root = Tk()
root.title("Absensi Perkuliahan")
root.geometry('800x600')
root.resizable(width=False, height=False)

# GUI elements
judul = Label(root, text='Absensi Perkuliahan', font=('Helvetica', 20, 'bold'))
judul.pack(pady=20)

matkul_label = Label(root, text='Mata kuliah:')
matkul_label.pack(anchor='w', padx=20)
matkul_entry = Entry(root)
matkul_entry.pack(anchor='w', padx=20)

tanggal_label = Label(root, text='Tanggal perkuliahan:')
tanggal_label.pack(anchor='w', padx=20)
tanggal_entry = Entry(root)
tanggal_entry.pack(anchor='w', padx=20)

nama_label = Label(root, text='Nama:')
nama_label.pack(anchor='w', padx=20)
nama_entry = Entry(root)
nama_entry.pack(anchor='w', padx=20)

NIM_label = Label(root, text='NIM:')
NIM_label.pack(anchor='w', padx=20)
NIM_entry = Entry(root)
NIM_entry.pack(anchor='w', padx=20)

jurusan_label = Label(root, text='Jurusan:')
jurusan_label.pack(anchor='w', padx=20)
jurusan_entry = Entry(root)
jurusan_entry.pack(anchor='w', padx=20)

informasi = Label(root, text='Klik Insert untuk semua mahasiswa, kemudian klik Save jika semua telah diabsen.')
informasi.pack(anchor='w', padx=20, pady=20)

# Frame for buttons
button_frame = Frame(root)
button_frame.pack(pady=20)

insert_button = Button(button_frame, text='Insert', command=insert_data, width=10)
insert_button.pack(side=LEFT, padx=10)

save_button = Button(button_frame, text='Save', command=save_data, width=10)
save_button.pack(side=LEFT, padx=10)

new_button = Button(button_frame, text='Create New', command=create_new_data, width=10)
new_button.pack(side=LEFT, padx=10)

calculate_button = Button(button_frame, text='Calculate', command=calculate_percentage, width=10)
calculate_button.pack(side=LEFT, padx=10)

# Excel styles
font_bold = Font(bold=True)
border = Border(left=Side(border_style='thin'), right=Side(border_style='thin'), top=Side(border_style='thin'), bottom=Side(border_style='thin'))
alignment_center = Alignment(horizontal='center', vertical='center')

# Create workbook and set initial values
workbook = Workbook()
sheet = workbook.active
num = 0  # Counter for number of students

# Start the main event loop
root.mainloop()
