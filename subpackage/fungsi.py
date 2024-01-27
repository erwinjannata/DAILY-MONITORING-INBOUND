import xlwings as xl
from tkinter.messagebox import showinfo
import re
import os

cabang_col1 = ['F', 'S', 'AF', 'AS', 'BF', 'BS', 'CF', 'CS',
               'DF', 'DS', 'EF', 'ES', 'FF', 'FS', 'GF', 'GS']

cabang_col2 = ['R', 'AE', 'AR', 'BE', 'BR', 'CE', 'CR',
               'DE', 'DR', 'EE', 'ER', 'FE', 'FR', 'GE', 'GR', 'HE']


def gabung_cabang(file_data, file_report, tgl, saved_as):
    tanggal = int(tgl.split('/')[1])

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

            for idx, cell_column1 in enumerate(cabang_col1[0:tanggal]):
                first_row = target_worksheet.range(f"{cell_column1}4")
                last_row = target_worksheet.range(f"{cell_column1}{max_row}")

                if first_row.value is None and target_worksheet.range(f"{cell_column1}{4 + merged}").value is None:
                    cell_row = 4
                else:
                    cell_row = int(re.findall(
                        r'\d+', (last_row.end('up').address))[0]) + 1

                if source_worksheet.range(f"{cell_column1}4").value is None:
                    col1 = int(re.findall(r'\d+', source_worksheet.range(f"{cell_column1}4").end('down').get_address(
                        row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])
                else:
                    col1 = 4

                col2 = int(re.findall(r'\d+', source_worksheet.range(f'{cell_column1}{col1}').end('down').get_address(
                    row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])

                # proses copy
                if (cell_row <= max_row) and target_worksheet.range(f"{cell_column1}{cell_row}").value is None:
                    source_worksheet.range(
                        f"{cell_column1}{col1}:{cabang_col2[idx]}{col2}").expand("down").copy()
                    target_worksheet.range(f"{cell_column1}{cell_row}").expand(
                        "table").paste(paste="values")

                    target_worksheet.api.Application.CutCopyMode = False
                else:
                    continue

            if target_worksheet.range("D4").value is None and target_worksheet.range(f"C{4 + merged}").value is None:
                cell_row = 4
            else:
                cell_row = int(re.findall(
                    r'\d+', (target_worksheet.range(f"D{max_row}").end('up').address))[0]) + 1

            col1 = int(re.findall(r'\d+', source_worksheet.range("F4").end('down').get_address(
                row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])

            if (cell_row < max_row) and target_worksheet.range(f"D{cell_row}").value is None:
                source_worksheet.range(
                    f"D{col1}:E{max_row}").expand("down").copy()
                target_worksheet.range(f"D{cell_row}").expand(
                    "table").paste(paste="values")

        target_workbook.save(saved_as)
        source_workbook.close()
        target_workbook.close()
        app.quit()
        showinfo(title="Message", message="Proses selesai")
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


customer_col1 = ['E', 'P', 'AA', 'AL', 'AW', 'BH', 'BS', 'CD', 'CO', 'CZ']

customer_col2 = ['O', 'Z', 'AK', 'AV', 'BG', 'BR', 'CC', 'CN', 'CY', 'DJ']


def gabung_customer(file_data, file_report, tgl, saved_as):
    tanggal = int(tgl.split('/')[1])

    if tanggal > 11:
        tanggal = 11

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

            for idx, cell_column1 in enumerate(customer_col1[0:tanggal]):
                first_row = target_worksheet.range(f"{cell_column1}5")
                last_row = target_worksheet.range(f"{cell_column1}{max_row}")

                if first_row.value is None and target_worksheet.range(f"{cell_column1}{5 + merged}").value is None:
                    cell_row = 5
                else:
                    cell_row = int(re.findall(
                        r'\d+', (last_row.end('up').address))[0])+1

                if source_worksheet.range(f"{cell_column1}5").value is None:
                    col1 = int(re.findall(r'\d+', source_worksheet.range(f"{cell_column1}5").end('down').get_address(
                        row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])
                else:
                    col1 = 5

                col2 = int(re.findall(r'\d+', source_worksheet.range(f'{cell_column1}{col1}').end('down').get_address(
                    row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])

                # proses copy
                if (cell_row < max_row) and target_worksheet.range(f"{cell_column1}{cell_row}").value is None:
                    source_worksheet.range(
                        f"{cell_column1}{col1}:{customer_col2[idx]}{col2}").expand("down").copy()
                    target_worksheet.range(f"{cell_column1}{cell_row}").expand(
                        "table").paste(paste="values")

                    target_worksheet.api.Application.CutCopyMode = False
                else:
                    continue

            if target_worksheet.range("C5").value is None and target_worksheet.range(f"C{5 + merged}").value is None:
                cell_row = 5
            else:
                cell_row = int(re.findall(
                    r'\d+', (target_worksheet.range(f"C{max_row}").end('up').address))[0])+1

            col1 = int(re.findall(r'\d+', source_worksheet.range("E5").end('down').get_address(
                row_absolute=False, column_absolute=False, include_sheetname=False, external=False))[0])

            if (cell_row < max_row) and target_worksheet.range(f"C{cell_row}").value is None:
                source_worksheet.range(
                    f"C{col1}:D{max_row}").expand("table").copy()
                target_worksheet.range(f"C{cell_row}").expand(
                    "table").paste(paste="values")

        if tanggal >= 11:
            for i in range(5, 10):
                source_worksheet = source_workbook.sheets[i]
                target_worksheet = target_workbook.sheets[i-5]

                if target_worksheet.range("DK5").value is None and target_worksheet.range(f"DK{5 + merged}").value is None:
                    cell_row = 5
                else:
                    cell_row = int(re.findall(
                        r'\d+', (target_worksheet.range(f"DK{max_row}").end('up').address))[0])+1

                if (cell_row < max_row) and target_worksheet.range(f"DK{cell_row}").value is None:
                    if i == 9:
                        source_worksheet.range("C5").expand('table').copy()
                    else:
                        source_worksheet.range("C4").expand('table').copy()
                    target_worksheet.range(f"DK{cell_row}").expand(
                        "table").paste(paste="values")

        target_workbook.save(saved_as)
        source_workbook.close()
        target_workbook.close()
        app.quit()
        showinfo(title="Message", message="Proses selesai")
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
