import xlwings as xl
from tkinter.messagebox import showinfo
import re
import os
import datetime as dt

cabang_col1 = ['D', 'S', 'AF', 'AS', 'BF', 'BS', 'CF', 'CS',
               'DF', 'DS', 'EF', 'ES', 'FF', 'FS', 'GF', 'GS']

cabang_col2 = ['R', 'AE', 'AR', 'BE', 'BR', 'CE', 'CR',
               'DE', 'DR', 'EE', 'ER', 'FE', 'FR', 'GE', 'GR', 'HE']


def gabung_cabang(file_data, file_report, tgl, saved_as, over_month):
    tanggal = int(tgl.split('/')[1])
    real_date = tanggal

    if tanggal > 16:
        tanggal = 16

    app = xl.App(visible=False)
    try:
        os.rename(file_report, file_report)

        global source_workbook, target_workbook
        source_workbook = xl.Book(file_data)
        target_workbook = xl.Book(file_report)

        for i in range(0, 12):
            source_worksheet = source_workbook.sheets[i+1]
            target_worksheet = target_workbook.sheets[i]

            global max_row, merged
            max_row = int(re.findall(
                r'\d+', (target_worksheet.range("C4").end("down").address))[0])

            merged = target_worksheet.range("A4").merge_area.count

            if over_month == 1:
                tanggal = 16
                cell_row = max_row + (((real_date - 1) * merged) + 1)
            else:
                cell_row = 4 + (merged * (real_date - 1))

            for idx, cell_column1 in enumerate(cabang_col1[0:tanggal]):
                # Lacak data yang akan di copy
                if source_worksheet.range(f"{cell_column1}4").value is None:
                    col1 = int(re.findall(r'\d+', source_worksheet.range(f"{cell_column1}4").end('down').get_address(
                        row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])
                else:
                    if idx == 0:
                        col1 = int(re.findall(
                            r'\d+', (source_worksheet.range("C4").end("down").address))[0]) - (merged - 1)
                    else:
                        col1 = 4

                if idx == 0:
                    col2 = int(re.findall(
                        r'\d+', (source_worksheet.range("C4").end("down").address))[0])
                else:
                    col2 = int(re.findall(r'\d+', source_worksheet.range(f'{cell_column1}{col1}').end('down').get_address(
                        row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])

                # proses copy
                if cell_row < max_row:
                    source_worksheet.range(
                        f"{cell_column1}{col1}:{cabang_col2[idx]}{col2}").expand("down").copy()
                    target_worksheet.range(f"{cell_column1}{cell_row}").expand(
                        "table").paste(paste="values")

                    target_worksheet.api.Application.CutCopyMode = False
                    cell_row -= merged
                else:
                    cell_row -= merged
                    continue

        target_workbook.save(saved_as)
        source_workbook.close()
        target_workbook.close()
        app.quit()
        showinfo(title="Message",
                 message=f"Proses selesai \n Hasil disimpan di: \n {saved_as}")
    except OSError:
        app.quit()
        showinfo(title="Message",
                 message="File excel sedang dibuka / digunakan oleh proses lain.")
    except Exception as e:
        source_workbook.close()
        target_workbook.close()
        app.quit()
        showinfo(title="Message",
                 message="Program mengalami masalah, silahkan hubungi tim IT.")
        print(e)


customer_col1 = ['C', 'P', 'AA', 'AL', 'AW', 'BH', 'BS', 'CD', 'CO', 'CZ']

customer_col2 = ['O', 'Z', 'AK', 'AV', 'BG', 'BR', 'CC', 'CN', 'CY', 'DJ']


def gabung_customer(file_data, file_report, tgl, saved_as, over_month):
    tanggal = int(tgl.split('/')[1])
    real_date = tanggal

    if tanggal > 15:
        tanggal = 10
    elif tanggal <= 15 and tanggal > 10:
        tanggal = 9
    elif tanggal <= 10 and tanggal > 7:
        tanggal = 8

    app = xl.App(visible=False)
    try:
        os.rename(file_report, file_report)

        global source_workbook, target_workbook
        source_workbook = xl.Book(file_data)
        target_workbook = xl.Book(file_report)

        for i in range(0, 5):
            source_worksheet = source_workbook.sheets[i]
            target_worksheet = target_workbook.sheets[i]

            global max_row, merged
            max_row = int(re.findall(
                r'\d+', (target_worksheet.range("B5").end("down").address))[0])

            merged = target_worksheet.range("A5").merge_area.count

            if over_month == 1:
                tanggal = 10
                cell_row = max_row + (((real_date - 1) * merged) + 1)
            else:
                cell_row = 5 + (merged * (real_date - 1))

            for idx, cell_column1 in enumerate(customer_col1[0:tanggal]):
                # Lacak data yang akan di copy
                if source_worksheet.range(f"{cell_column1}5").value is None:
                    col1 = int(re.findall(r'\d+', source_worksheet.range(f"{cell_column1}5").end('down').get_address(
                        row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])
                else:
                    if idx == 0:
                        col1 = int(re.findall(
                            r'\d+', (source_worksheet.range("C5").end("down").address))[0]) - (merged - 1)
                    else:
                        col1 = 5

                if idx == 0:
                    col2 = int(re.findall(
                        r'\d+', (source_worksheet.range("C5").end("down").address))[0])
                else:
                    col2 = int(re.findall(r'\d+', source_worksheet.range(f'{cell_column1}{col1}').end('down').get_address(
                        row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])

                # proses copy
                if cell_row <= max_row:
                    source_worksheet.range(
                        f"{cell_column1}{col1}:{customer_col2[idx]}{col2}").expand("down").copy()
                    target_worksheet.range(f"{cell_column1}{cell_row}").expand(
                        "table").paste(paste="values")

                    target_worksheet.api.Application.CutCopyMode = False
                    if idx < 7:
                        cell_row -= merged
                    elif idx == 7:
                        cell_row -= (merged * 3)
                    elif idx == 8:
                        cell_row -= (merged * 5)
                else:
                    if idx < 7:
                        cell_row -= merged
                    elif idx == 7:
                        cell_row -= (merged * 3)
                    elif idx == 8:
                        cell_row -= (merged * 5)
                    continue

        if real_date >= 17 or over_month == 1:
            for i in range(5, 10):
                source_worksheet = source_workbook.sheets[i]
                target_worksheet = target_workbook.sheets[i-5]

                cell_row = (real_date - 16) * merged

                if over_month == 1:
                    cell_row = ((31 + real_date) - 16) * merged

                if cell_row <= max_row:
                    if i == 9:
                        source_worksheet.range("C5").expand('table').copy()
                    else:
                        source_worksheet.range("C4").expand('table').copy()
                    target_worksheet.range(f"DK{cell_row}").expand(
                        "table").paste(paste="values")

                    target_worksheet.api.Application.CutCopyMode = False

        target_workbook.save(saved_as)
        source_workbook.close()
        target_workbook.close()
        showinfo(title="Message",
                 message=f"Proses selesai \n Hasil disimpan di: \n {saved_as}")
    except OSError:
        app.quit()
        showinfo(title="Message",
                 message="File excel sedang dibuka / digunakan oleh proses lain.")
    except Exception as e:
        source_workbook.close()
        target_workbook.close()
        app.quit()
        showinfo(title="Message",
                 message="Program mengalami masalah, silahkan hubungi tim IT.")
        print(e)
