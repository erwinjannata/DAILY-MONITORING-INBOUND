import re
import os
import pandas as pd
import xlwings as xl
from tkinter.messagebox import showinfo
from subpackage.grouping import grouping_daily_monitor
from subpackage.counting import counting_cabang, counting_customer


def gabung_cabang(file_data, file_report, tgl, saved_as, over_month, is_grouped):
    app = xl.App(visible=False)
    target_workbook = xl.Book(file_report)

    try:
        if is_grouped == 0:
            datas = grouping_daily_monitor(
                file_data=file_data, save_grouping=False, saved_as='', tanggal=tgl)
        else:
            datas = pd.read_excel(file_data)

        real_date = int(tgl.split('/')[1])

        # Do process for every sheet in File Report
        for i in range(0, 12):
            target_worksheet = target_workbook.sheets[i]

            # Load Processed Data
            processed_data = counting_cabang(
                df_cnote=datas, sheet=target_worksheet, date=tgl, nowIndex=i)
            # Load data for H+0 - H+X
            count_data = processed_data[1]
            # Load data for COD
            cod_data = processed_data[0]

            # Analyze File Report spesification
            global max_row, merged_row, merged_col
            max_row = int(re.findall(
                r'\d+', (target_worksheet.range("C4").end("down").address))[0])
            merged_row = target_worksheet.range("A4").merge_area.count
            merged_col = target_worksheet.range("F1").merge_area.count
            jumlah_zona = target_worksheet.range("B4").merge_area.count

            # Condition to give space for every customer
            if jumlah_zona > 1:
                jumlah_zona -= 1
                spasi = jumlah_zona + 1
            else:
                spasi = 1

            # Condition if writing to previous month report
            if over_month == 1:
                cell_row = max_row + (((real_date - 1) * merged_row))
            else:
                cell_row = (4 + (merged_row * (real_date - 1))) - 1

            # Fill data for COD on H+0
            if cell_row >= 3 and cell_row < max_row:
                # Paste value for every Customer
                for cust in range(0, 5):
                    sum_cod_ship = 0
                    sum_cod_amount = 0

                    # Paste value for every Zona
                    for zona in range(0, jumlah_zona):
                        # Total Shipment COD
                        target_worksheet[(
                            (cell_row + zona) + (cust * spasi)), 3].value = cod_data[0][cust][zona][0]
                        # Total Nominal COD
                        target_worksheet[(
                            (cell_row + zona) + (cust * spasi)), 4].value = cod_data[0][cust][zona][1]

                        # Add sum of values
                        sum_cod_ship += cod_data[0][cust][zona][0]
                        sum_cod_amount += cod_data[0][cust][zona][1]

                    # Fill value for Sum
                    if jumlah_zona > 1:
                        addition = cust

                        # Sum COD Shipment
                        target_worksheet[(
                            cell_row + (jumlah_zona * (cust + 1)) + addition), 3].value = sum_cod_ship

                        # Sum COD Amount
                        target_worksheet[(
                            cell_row + (jumlah_zona * (cust + 1)) + addition), 4].value = sum_cod_amount

            # Fill data for H+0 - H+x
            for hari in range(0, 16):
                # Condition to prevent leaking
                if cell_row >= 3 and cell_row < max_row:
                    # Paste value for every Customer
                    for cust in range(0, 5):
                        # Record sum of values
                        sum_total_cnote = 0
                        sum_unrunsheet = 0
                        sum_sukses = 0
                        sum_cr = 0
                        sum_undel = 0
                        sum_unstatus = 0
                        sum_wh1 = 0
                        sum_irreg = 0
                        sum_return = 0

                        # Paste value for every Zona
                        for zona in range(0, jumlah_zona):

                            # Value for every Categories
                            # Total Cnote
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 5 + (merged_col * hari)].value = count_data[hari][cust][zona][0]
                            # Unrunsheet
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 6 + (merged_col * hari)].value = count_data[hari][cust][zona][1]
                            # Sukses
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 7 + (merged_col * hari)].value = count_data[hari][cust][zona][2]
                            # CR
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 8 + (merged_col * hari)].value = count_data[hari][cust][zona][3]
                            # Undel
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 9 + (merged_col * hari)].value = count_data[hari][cust][zona][4]
                            # Unstatus
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 10 + (merged_col * hari)].value = count_data[hari][cust][zona][5]
                            # WH1
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 11 + (merged_col * hari)].value = count_data[hari][cust][zona][6]
                            # Irregularity
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 12 + (merged_col * hari)].value = count_data[hari][cust][zona][7]
                            # Return
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 13 + (merged_col * hari)].value = count_data[hari][cust][zona][8]

                            # Add sum of values
                            sum_total_cnote += count_data[hari][cust][zona][0]
                            sum_unrunsheet += count_data[hari][cust][zona][1]
                            sum_sukses += count_data[hari][cust][zona][2]
                            sum_cr += count_data[hari][cust][zona][3]
                            sum_undel += count_data[hari][cust][zona][4]
                            sum_unstatus += count_data[hari][cust][zona][5]
                            sum_wh1 += count_data[hari][cust][zona][6]
                            sum_irreg += count_data[hari][cust][zona][7]
                            sum_return += count_data[hari][cust][zona][8]

                            # Percentage (%) of every categories
                            # % Sukses
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 14 + (merged_col * hari)].value = (count_data[hari][cust][zona][2] / count_data[hari][cust][zona][0]) if count_data[hari][cust][zona][0] != 0 else 0
                            # % Unrunsheet
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 15 + (merged_col * hari)].value = (count_data[hari][cust][zona][1] / count_data[hari][cust][zona][0]) if count_data[hari][cust][zona][0] != 0 else 0
                            # % Return
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 16 + (merged_col * hari)].value = (count_data[hari][cust][zona][8] / count_data[hari][cust][zona][0]) if count_data[hari][cust][zona][0] != 0 else 0
                            # % Failure
                            target_worksheet[(
                                (cell_row + zona) + (cust * spasi)), 17 + (merged_col * hari)].value = ((count_data[hari][cust][zona][1] + count_data[hari][cust][zona][3] + count_data[hari][cust][zona][4] + count_data[hari][cust][zona][5] + count_data[hari][cust][zona][6] + count_data[hari][cust][zona][7] + count_data[hari][cust][zona][8]) / count_data[hari][cust][zona][0]) if count_data[hari][cust][zona][0] != 0 else 0

                        if jumlah_zona > 1:
                            addition = cust

                            # Count Sum value of all Zona
                            # Sum Total Cnote
                            target_worksheet[(cell_row + (jumlah_zona * (cust + 1)) + addition), 5 + (
                                merged_col * hari)].value = sum_total_cnote
                            # Sum Unrunsheet
                            target_worksheet[(cell_row + (jumlah_zona * (cust + 1)) + addition),
                                             6 + (merged_col * hari)].value = sum_unrunsheet
                            # Sum Sukses
                            target_worksheet[(
                                cell_row + (jumlah_zona * (cust + 1)) + addition), 7 + (merged_col * hari)].value = sum_sukses
                            # Sum CR
                            target_worksheet[(
                                cell_row + (jumlah_zona * (cust + 1)) + addition), 8 + (merged_col * hari)].value = sum_cr
                            # Sum Undel
                            target_worksheet[(
                                cell_row + (jumlah_zona * (cust + 1)) + addition), 9 + (merged_col * hari)].value = sum_undel
                            # Sum Unstatus
                            target_worksheet[(cell_row + (jumlah_zona * (cust + 1)) + addition),
                                             10 + (merged_col * hari)].value = sum_unstatus
                            # Sum WH1
                            target_worksheet[(
                                cell_row + (jumlah_zona * (cust + 1)) + addition), 11 + (merged_col * hari)].value = sum_wh1
                            # Sum Irreg
                            target_worksheet[(
                                cell_row + (jumlah_zona * (cust + 1)) + addition), 12 + (merged_col * hari)].value = sum_irreg
                            # Sum Return
                            target_worksheet[(
                                cell_row + (jumlah_zona * (cust + 1)) + addition), 13 + (merged_col * hari)].value = sum_return

                            # Count % Sum value of all Zona
                            # % Sum Sukses
                            target_worksheet[(
                                cell_row + (jumlah_zona * (cust + 1)) + addition), 14 + (merged_col * hari)].value = sum_sukses / sum_total_cnote if sum_total_cnote != 0 else 0
                            # % Sum Unrunsheet
                            target_worksheet[(
                                cell_row + (jumlah_zona * (cust + 1)) + addition), 15 + (merged_col * hari)].value = sum_unrunsheet / sum_total_cnote if sum_total_cnote != 0 else 0
                            # % Sum Return
                            target_worksheet[(
                                cell_row + (jumlah_zona * (cust + 1)) + addition), 16 + (merged_col * hari)].value = sum_return / sum_total_cnote if sum_total_cnote != 0 else 0
                            # % Sum Failed
                            target_worksheet[(
                                cell_row + (jumlah_zona * (cust + 1)) + addition), 17 + (merged_col * hari)].value = (sum_unrunsheet + sum_cr + sum_undel + sum_unstatus + sum_wh1 + sum_irreg + sum_return) / sum_total_cnote if sum_total_cnote != 0 else 0
                    cell_row -= merged_row
                else:
                    cell_row -= merged_row
                    continue

        target_workbook.save(saved_as)
        target_workbook.close()
        app.quit()
        showinfo(title="Message",
                 message=f"Proses selesai \n Hasil disimpan di: \n {saved_as}")
    except Exception as e:
        target_workbook.close()
        app.quit()
        showinfo(title="Message",
                 message="Program mengalami masalah, silahkan hubungi tim IT.")
        print(e)


def gabung_customer(file_data, file_report, tgl, saved_as, over_month, is_grouped):
    app = xl.App(visible=False)
    target_workbook = xl.Book(file_report)

    try:
        real_date = int(tgl.split('/')[1])

        processed_data = counting_customer(
            file_data=file_data, date=tgl, is_grouped=is_grouped)
        count_data = processed_data[1]
        cod_data = processed_data[0]

        # Fill data for every sheet
        for i in range(0, 5):
            target_worksheet = target_workbook.sheets[i]

            # Analyze File Report spesification
            global max_row, merged_row, merged_col
            max_row = int(re.findall(
                r'\d+', (target_worksheet.range("B5").end("down").address))[0])
            merged_row = target_worksheet.range("A5").merge_area.count
            merged_col = target_worksheet.range("E2").merge_area.count

            # Condition if writing to previous month report
            if over_month == 1:
                cell_row = max_row + (merged_row * (real_date - 1))
            else:
                cell_row = (merged_row * real_date) - 1

            # Fill data for COD on H+0
            if cell_row >= 4 and cell_row < max_row:
                sum_cod_ship = 0
                sum_cod_amount = 0
                # Paste value for every Zona
                for zona in range(0, 4):
                    # Total Shipment COD
                    target_worksheet[(cell_row + zona),
                                     2].value = cod_data[0][i][zona][0]
                    # Total Nominal COD
                    target_worksheet[(cell_row + zona),
                                     3].value = cod_data[0][i][zona][1]
                    # Add sum of values
                    sum_cod_ship += cod_data[0][i][zona][0]
                    sum_cod_amount += cod_data[0][i][zona][1]
                # Fill value for Sum
                # Sum COD Shipment
                target_worksheet[(cell_row + 4), 2].value = sum_cod_ship
                # Sum COD Amount
                target_worksheet[(cell_row + 4), 3].value = sum_cod_amount

            # Fill H+0 - H+7, H+10, H+15
            for hari in range(0, 11):
                hari_ref = hari
                if hari == 8:
                    hari_ref += 2
                elif hari == 9 or hari == 10:
                    hari_ref += 6

                # Record sum of values
                sum_total_cnote = 0
                sum_unrunsheet = 0
                sum_sukses = 0
                sum_cr = 0
                sum_undel = 0
                sum_unstatus = 0
                sum_wh1 = 0
                sum_irreg = 0
                sum_return = 0
                sum_breach = 0

                per_sukses_col_position = 11 if hari != 10 else 13
                per_unrunsheet_col_position = 12 if hari != 10 else 14
                per_return_col_position = 13 if hari != 10 else 15
                per_failed_col_position = 14 if hari != 10 else 18

                unstatus_col_position = 9 if hari != 10 else 10
                return_col_position = 10 if hari != 10 else 11

                if cell_row >= 4 and cell_row < max_row:
                    for zona in range(0, 4):
                        # Value for every Categories
                        # Total Cnote
                        target_worksheet[
                            (cell_row + zona), 4 + (merged_col * hari)].value = count_data[hari][i][zona][0]

                        # Unrunsheet
                        target_worksheet[
                            (cell_row + zona), 5 + (merged_col * hari)].value = count_data[hari][i][zona][1]

                        # Sukses
                        target_worksheet[
                            (cell_row + zona), 6 + (merged_col * hari)].value = count_data[hari][i][zona][2]

                        # CR
                        target_worksheet[
                            (cell_row + zona), 7 + (merged_col * hari)].value = count_data[hari][i][zona][3]

                        # Undel
                        target_worksheet[
                            (cell_row + zona), 8 + (merged_col * hari)].value = count_data[hari][i][zona][4]

                        # Irregularity (Fill only for H>15)
                        if hari == 10:
                            target_worksheet[(
                                cell_row + zona), 9 + (merged_col * hari)].value = count_data[hari][i][zona][7]

                        # Unstatus
                        target_worksheet[
                            (cell_row + zona), unstatus_col_position + (merged_col * hari)].value = count_data[hari][i][zona][5]

                        # Return
                        target_worksheet[
                            (cell_row + zona), return_col_position + (merged_col * hari)].value = count_data[hari][i][zona][6]

                        # Breach (Fill only for H>15)
                        if hari == 10:
                            target_worksheet[(
                                cell_row + zona), 12 + (merged_col * hari)].value = count_data[hari][i][zona][8]

                        # Add sum of values
                        sum_total_cnote += count_data[hari][i][zona][0]
                        sum_unrunsheet += count_data[hari][i][zona][1]
                        sum_sukses += count_data[hari][i][zona][2]
                        sum_cr += count_data[hari][i][zona][3]
                        sum_undel += count_data[hari][i][zona][4]
                        sum_unstatus += count_data[hari][i][zona][5]
                        sum_irreg += count_data[hari][i][zona][7]
                        sum_return += count_data[hari][i][zona][6]
                        sum_breach += count_data[hari][i][zona][8]

                        # Percentage (%) of every categories
                        # % Sukses
                        target_worksheet[(cell_row + zona), per_sukses_col_position + (merged_col * hari)].value = (
                            count_data[hari][i][zona][2] / count_data[hari][i][zona][0]) if count_data[hari][i][zona][0] != 0 else 0
                        # % Unrunsheet
                        target_worksheet[(cell_row + zona), per_unrunsheet_col_position + (merged_col * hari)].value = (
                            count_data[hari][i][zona][1] / count_data[hari][i][zona][0]) if count_data[hari][i][zona][0] != 0 else 0
                        # % Return
                        target_worksheet[(cell_row + zona), per_return_col_position + (merged_col * hari)].value = (
                            count_data[hari][i][zona][6] / count_data[hari][i][zona][0]) if count_data[hari][i][zona][0] != 0 else 0
                        # % Irregularity (Fill only on H>15)
                        if hari == 10:
                            target_worksheet[(cell_row + zona), 16 + (merged_col * hari)].value = (
                                count_data[hari][i][zona][7] / count_data[hari][i][zona][0]) if count_data[hari][i][zona][0] != 0 else 0
                        # % Breach (Fill only on H>15)
                        if hari == 10:
                            target_worksheet[(cell_row + zona), 17 + (merged_col * hari)].value = (
                                count_data[hari][i][zona][8] / count_data[hari][i][zona][0]) if count_data[hari][i][zona][0] != 0 else 0
                        # % Failure
                        failure = (
                            count_data[hari][i][zona][1] +
                            count_data[hari][i][zona][3] +
                            count_data[hari][i][zona][4] +
                            count_data[hari][i][zona][5] +
                            count_data[hari][i][zona][6])
                        failure_final = (
                            failure + count_data[hari][i][zona][7] + count_data[hari][i][zona][8]) if hari == 10 else failure
                        perc_failed = (
                            failure_final / count_data[hari][i][zona][0]) if count_data[hari][i][zona][0] != 0 else 0
                        target_worksheet[(
                            cell_row + zona), per_failed_col_position + (merged_col * hari)].value = perc_failed

                    # Count Sum value of all Zona
                    # Sum Total Cnote
                    target_worksheet[(cell_row + 4), 4 + (
                        merged_col * hari)].value = sum_total_cnote
                    # Sum Unrunsheet
                    target_worksheet[(cell_row + 4),
                                     5 + (merged_col * hari)].value = sum_unrunsheet
                    # Sum Sukses
                    target_worksheet[(cell_row + 4), 6 +
                                     (merged_col * hari)].value = sum_sukses
                    # Sum CR
                    target_worksheet[(cell_row + 4), 7 +
                                     (merged_col * hari)].value = sum_cr
                    # Sum Undel
                    target_worksheet[(cell_row + 4), 8 +
                                     (merged_col * hari)].value = sum_undel
                    # Sum Irreg (Fill only on H>15)
                    if hari == 10:
                        target_worksheet[(cell_row + 4), 9 +
                                         (merged_col * hari)].value = sum_irreg
                    # Sum Unstatus
                    target_worksheet[(cell_row + 4),
                                     unstatus_col_position + (merged_col * hari)].value = sum_unstatus
                    # Sum Return
                    target_worksheet[(cell_row + 4), return_col_position +
                                     (merged_col * hari)].value = sum_return

                    # Count % Sum value of all Zona
                    # % Sum Sukses
                    target_worksheet[(cell_row + 4), per_sukses_col_position + (merged_col * hari)].value = (
                        sum_sukses / sum_total_cnote) if sum_total_cnote != 0 else 0
                    # % Sum Unrunsheet
                    target_worksheet[(cell_row + 4), per_unrunsheet_col_position + (merged_col * hari)].value = (
                        sum_unrunsheet / sum_total_cnote) if sum_total_cnote != 0 else 0
                    # % Sum Return
                    target_worksheet[(cell_row + 4), per_return_col_position + (merged_col * hari)].value = (
                        sum_return / sum_total_cnote) if sum_total_cnote != 0 else 0
                    # % Sum Irreg
                    if hari == 10:
                        target_worksheet[(cell_row + 4), 16 +
                                         (merged_col * hari)].value = (sum_irreg / sum_total_cnote) if sum_total_cnote != 0 else 0
                    # % Sum Breach
                    if hari == 10:
                        target_worksheet[(cell_row + 4), 17 +
                                         (merged_col * hari)].value = (sum_breach / sum_total_cnote) if sum_total_cnote != 0 else 0
                    # % Sum Failed
                    sum_failed = (sum_unrunsheet + sum_cr +
                                  sum_undel + sum_unstatus + sum_return)
                    sum_failed_final = (
                        sum_failed + sum_irreg + sum_breach) if hari == 10 else sum_failed
                    target_worksheet[(cell_row + 4), per_failed_col_position + (merged_col * hari)].value = (
                        sum_failed_final / sum_total_cnote) if sum_total_cnote != 0 else 0

                    if hari == 7:
                        cell_row -= (merged_row * 3)
                    elif hari == 8:
                        cell_row -= (merged_row * 5)
                    else:
                        cell_row -= merged_row
                else:
                    if hari == 7:
                        cell_row -= (merged_row * 3)
                    elif hari == 8:
                        cell_row -= (merged_row * 5)
                    else:
                        cell_row -= merged_row
                    continue

        target_workbook.save(saved_as)
        target_workbook.close()
        showinfo(title="Message",
                 message=f"Proses selesai \n Hasil disimpan di: \n {saved_as}")
    except Exception as e:
        target_workbook.close()
        app.quit()
        showinfo(title="Message",
                 message="Program mengalami masalah, silahkan hubungi tim IT.")
        print(e)
