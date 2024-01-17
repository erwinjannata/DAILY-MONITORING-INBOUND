import xlwings as xl
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter.messagebox import showinfo
from tkcalendar import DateEntry
import threading
import re
import os

root = tk.Tk()
root.configure(bg="white")
root.geometry("350x360")
root.title("Penggabung Data Monitoring")

file_data = ""
file_report = ""

file_data_name = tk.StringVar()
file_report_name = tk.StringVar()

column_data1 = ['F', 'S', 'AF', 'AS', 'BF', 'BS', 'CF', 'CS',
                'DF', 'DS', 'EF', 'ES', 'FF', 'FS', 'GF', 'GS']

column_data2 = ['R', 'AE', 'AR', 'BE', 'BR', 'CE', 'CR',
                'DE', 'DR', 'EE', 'ER', 'FE', 'FR', 'GE', 'GR', 'HE']


def load_data():
    global file_data
    file_data = filedialog.askopenfilename(filetypes=(
        ("Excel Workbook", "*.xlsx"),
        ("All Files", "*.*"),
    ))
    file_data_name.set(os.path.split(file_data)[1])


def load_master_report():
    global file_report
    file_report = filedialog.askopenfilename(filetypes=(
        ("Excel Workbook", "*.xlsx"),
        ("All Files", "*.*"),
    ))
    file_report_name.set(os.path.split(file_report)[1])


def combine_process():
    date = calendar.get().split('/')
    tanggal = int(date[0])

    if tanggal > 16:
        tanggal = 16

    if file_data and file_report:
        app = xl.App(visible=False)
        source_workbook = xl.Book(file_data)
        target_workbook = xl.Book(file_report)

        try:
            for i in range(0, 12):
                source_worksheet = source_workbook.sheets[i+1]
                target_worksheet = target_workbook.sheets[i]

                global max_row
                max_row = int(re.findall(
                    r'\d+', (target_worksheet.range("C4").end("down").address))[0])

                for idx, cell_column1 in enumerate(column_data1[0:tanggal]):
                    cell = target_worksheet.range(f"{cell_column1}4")
                    if cell.value is None:
                        cell_row = 4
                    else:
                        cell_row = int(re.findall(
                            r'\d+', (cell.end('down').address))[0]) + 1

                    if source_worksheet.range(f"{cell_column1}4").value is None:
                        col1 = int(re.findall(r'\d+', source_worksheet.range(f"{cell_column1}4").end('down').get_address(
                            row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])
                    else:
                        col1 = 4

                    col2 = int(re.findall(r'\d+', source_worksheet.range(f'{cell_column1}{col1}').end('down').get_address(
                        row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])

                    # proses copy
                    if cell_row <= max_row:
                        source_worksheet.range(
                            f"{cell_column1}{col1}:{column_data2[idx]}{col2}").expand("down").copy()
                        target_worksheet.range(f"{cell_column1}{cell_row}").expand(
                            "table").paste(paste="values")

                        target_worksheet.api.Application.CutCopyMode = False
                    else:
                        continue

                if target_worksheet.range("D4").value is None:
                    cell_row = 4
                else:
                    cell_row = int(re.findall(
                        r'\d+', (target_worksheet.range("D4").end('down').address))[0]) + 1

                col1 = int(re.findall(r'\d+', source_worksheet.range("F4").end('down').get_address(
                    row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])

                if cell_row <= max_row:
                    source_worksheet.range(
                        f"D{col1}:E{max_row}").expand("down").copy()
                    target_worksheet.range(f"D{cell_row}").expand(
                        "table").paste(paste="values")

                target_workbook.save()

            source_workbook.close()
            target_workbook.close()
            app.quit()
            showinfo(title="Message",
                     message="Proses selesai")
        except Exception as e:
            source_workbook.close()
            target_workbook.close()
            app.quit()
            showinfo(title="Message",
                     message="Program mengalami masalah, silahkan hubungi tim IT")
    elif not file_data:
        showinfo(title="Message",
                 message="Tidak ada file data dipilih!")
    elif not file_report:
        showinfo(title="Message",
                 message="Tidak ada file master dipilih!")
    else:
        showinfo(title="Message",
                 message="Cek kembali excel yang dipilih!")
    btn1.state(['!disabled'])
    btn2.state(['!disabled'])
    combine_btn.state(['!disabled'])


def start_combine_thread(event):
    global combine_thread
    combine_thread = threading.Thread(target=combine_process)
    combine_thread.daemon = True
    progressbar.start()
    btn1.state(['disabled'])
    btn2.state(['disabled'])
    combine_btn.state(['disabled'])
    combine_thread.start()
    root.after(20, check_combine_thread)


def check_combine_thread():
    if combine_thread.is_alive():
        root.after(20, check_combine_thread)
    else:
        progressbar.stop()


# GUI
calendar = DateEntry(root, selectmode='day', locale='en_US',
                     date_pattern='dd/MM/yyyy', weekendbackground='white', weekendforeground='black')

label1 = ttk.Label(root, text="1. Tanggal Data", background="white").pack(
    fill="x", padx=10, pady=5)

calendar.pack(pady=10, padx=10, fill='both')

label2 = ttk.Label(root, text="2. Pilih file Excel Rumus & Data", background="white").pack(
    fill="x", padx=10, pady=5)

label_name1 = ttk.Label(root, textvariable=file_data_name, background="white").pack(
    fill="x", padx=10, pady=5)

btn1 = ttk.Button(root, text="Pilih File", command=load_data, state=tk.NORMAL)
btn1.pack(fill="x", padx=10, pady=5)

label3 = ttk.Label(root, text="3. Pilih file Excel Report", background="white").pack(
    fill="x", padx=10, pady=5)

label_name2 = ttk.Label(root, textvariable=file_report_name, background="white").pack(
    fill="x", padx=10, pady=5)

btn2 = ttk.Button(root, text="Pilih File",
                  command=load_master_report, state=tk.NORMAL)
btn2.pack(fill="x", padx=10, pady=5)

separator = ttk.Separator(root, orient='horizontal').pack(
    fill='x', pady=5, padx=10)

combine_btn = ttk.Button(root, text="Gabungkan",
                         command=lambda: start_combine_thread(None), state=tk.NORMAL)
combine_btn.pack(fill="x", padx=10, pady=10)

progressbar = ttk.Progressbar(root, mode='indeterminate')
progressbar.pack(fill='x', padx=10, pady=10)

root.mainloop()
