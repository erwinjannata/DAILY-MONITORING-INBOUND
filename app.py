import tkinter as tk
from tkinter import filedialog, ttk
from tkinter.messagebox import showinfo
from tkcalendar import DateEntry
import threading
import os
from subpackage.fungsi import gabung_cabang, gabung_customer

root = tk.Tk()
root.iconbitmap(r'D:\JNE\PROGRAM\DAILY MONITORING INBOUND\imgs\jne.ico')
root.configure(bg="white")
root.geometry("375x535")
root.title("Combine Data Monitoring")

file_data = ""
file_report = ""

file_data_name = tk.StringVar()
file_report_name = tk.StringVar()
mode = tk.StringVar()
over_month = tk.IntVar()
over_date = tk.IntVar()


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
    mode = combo_box.current()
    saved_as = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[
                                            ("Excel Workbook (.xlsx)", "*.xlsx")])
    date = calendar.get()

    if (os.path.exists(file_data) and os.path.exists(file_report)) and (saved_as and over_date.get() <= 31):
        progressbar.start()
        if mode == 0:
            gabung_cabang(file_data=file_data,
                          file_report=file_report, tgl=date, saved_as=saved_as, over_month=over_month.get(), tgl_over=over_date.get())
        elif mode == 1:
            gabung_customer(file_data=file_data,
                            file_report=file_report, tgl=date, saved_as=saved_as, over_month=over_month.get(), tgl_over=over_date.get())
        else:
            showinfo(title="Message",
                     message="Pilihan jenis report tidak valid")
    elif not file_data:
        showinfo(title="Message",
                 message="File rumus tidak ditemukan!")
    elif not file_report:
        showinfo(title="Message",
                 message="File report tidak ditemukan!")
    elif not saved_as:
        showinfo(title="Message",
                 message="Pilih lokasi penyimpanan file yang valid")
    elif mode == 0 and over_date.get() >= 16:
        showinfo(title="Message",
                 message="Tanggal bulan selanjutnya tidak valid!")
    elif over_date.get() >= 17:
        showinfo(title="Message",
                 message="Tanggal bulan selanjutnya tidak valid!")
    else:
        showinfo(title="Message",
                 message="Cek kembali excel yang dipilih!")
    combo_box.state(['!disabled'])
    calendar.state(['!disabled'])
    btn1.state(['!disabled'])
    btn2.state(['!disabled'])
    check_label.config(state="normal")
    over_tgl['state'] = 'disabled'
    over_month.set(0)
    over_date.set(0)
    combine_btn.state(['!disabled'])


def start_combine_thread(event):
    global combine_thread
    combine_thread = threading.Thread(target=combine_process)
    combine_thread.daemon = True
    combo_box.state(['disabled'])
    calendar.state(['disabled'])
    btn1.state(['disabled'])
    btn2.state(['disabled'])
    check_label.config(state="disabled")
    over_tgl['state'] = 'disabled'
    combine_btn.state(['disabled'])
    combine_thread.start()
    root.after(20, check_combine_thread)


def check_combine_thread():
    if combine_thread.is_alive():
        root.after(20, check_combine_thread)
    else:
        progressbar.stop()


def change_state():
    if over_tgl['state'] == 'disabled':
        over_tgl['state'] = 'normal'
        over_date.set(1)
    else:
        over_tgl['state'] = 'disabled'
        over_date.set(0)


# GUI
combo_label = ttk.Label(root, text="1. Jenis Report",
                        background="white", font="calibri 11 bold").pack(fill="x", padx=10, pady=5)
combo_box = ttk.Combobox(root, textvariable=mode)
combo_box['value'] = ('Daily Monitoring per Cabang',
                      'Daily Monitoring per Customer')
combo_box.pack(pady=10, padx=10, fill='both')
combo_box.current(0)

label1 = ttk.Label(root, text="2. Tanggal Data", background="white", font="calibri 11 bold").pack(
    fill="x", padx=10, pady=5)

calendar = DateEntry(root, selectmode='day', locale='en_US',
                     date_pattern='M/d/yyyy', weekendbackground='white', weekendforeground='black')
calendar.pack(pady=10, padx=10, fill='both')

check_label = tk.Checkbutton(
    root, text="Tarikan bulan selanjutnya", background="white", variable=over_month, onvalue=1, offvalue=0, command=change_state)
check_label.pack(pady=5, padx=10, anchor="w")

label4 = ttk.Label(root, text="3. Tanggal tarikan bulan selanjutnya **opsional", background="white", font="calibri 11 bold").pack(
    fill="x", padx=10, pady=5)

over_tgl = tk.Entry(root, textvariable=over_date, state='disabled')
over_tgl.pack(
    fill="x", padx=10, pady=5)

label2 = ttk.Label(root, text="4. File excel rumus", background="white", font="calibri 11 bold").pack(
    fill="x", padx=10, pady=5)

label_name1 = ttk.Label(root, textvariable=file_data_name, background="white").pack(
    fill="x", padx=10, pady=5)

btn1 = ttk.Button(root, text="Pilih File", command=load_data, state=tk.NORMAL)
btn1.pack(fill="x", padx=10, pady=5)

label3 = ttk.Label(root, text="5. File report terbaru", background="white", font="calibri 11 bold").pack(
    fill="x", padx=10, pady=5)

label_name2 = ttk.Label(root, textvariable=file_report_name, background="white").pack(
    fill="x", padx=10, pady=5)

btn2 = ttk.Button(root, text="Pilih File",
                  command=load_master_report, state=tk.NORMAL)
btn2.pack(fill="x", padx=10, pady=5)

separator = ttk.Separator(root, orient='horizontal').pack(
    fill='x', pady=5, padx=10)

combine_btn = ttk.Button(root, text="Proses",
                         command=lambda: start_combine_thread(None), state=tk.NORMAL)
combine_btn.pack(fill="x", padx=10, pady=10)

progressbar = ttk.Progressbar(root, mode='indeterminate', orient='horizontal')
progressbar.pack(fill='x', padx=10, pady=10)

root.mainloop()
