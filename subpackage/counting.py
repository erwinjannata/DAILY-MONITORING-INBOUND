import pandas as pd
from datetime import datetime, timedelta
from subpackage.grouping import grouping_daily_monitor

nama_cabang = ['MATARAM', 'BIMA', 'DOMPU', 'MANGGELEWA', 'SUMBAWA',
               'UTAN', 'ALAS', 'TALIWANG', 'LOTIM', 'PRAYA', 'LOBAR', 'TANJUNG']

nama_customer = ["LAZADA", "ORDIVO", "TOKOPEDIA"]


def counting_cabang(df_cnote, sheet, nowIndex, date):
    jumlah_zona = sheet.range("B4").merge_area.count
    if jumlah_zona > 1:
        jumlah_zona -= 1

    # Data COD H+0
    cod_all = []
    # ---------- ALL & COD ----------
    all_cod = []
    for zona in range(0, jumlah_zona):
        nama_zona = sheet[f'C{4 + zona}'].value
        # Total Ship COD
        ship_cod = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Hawb PCS'] == nama_zona) & (
            df_cnote['COD flag'] == 'Y') & (df_cnote['Manifest Bag No'] == datetime.strptime(date, '%m/%d/%Y'))])
        # Total Nominal COD
        cod_amount = df_cnote.loc[(df_cnote['Manifest Bag No'] == datetime.strptime(date, '%m/%d/%Y')) &
                                  (df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['COD flag'] == 'Y'), 'COD amount'].sum()
        all_cod.append([ship_cod, cod_amount])

    # ---------- LAZADA, ORDIVO, TOKOPEDIA ----------
    whole_cod = []
    for customer in nama_customer:
        cust_cod = []
        for zona in range(0, jumlah_zona):
            nama_zona = sheet[f'C{4 + zona}'].value
            # Total Ship COD
            ship_cod = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Hawb PCS'] == nama_zona) & (
                df_cnote['COD flag'] == 'Y') & (df_cnote['Hawb Customer Name'] == customer) & (df_cnote['Manifest Bag No'] == datetime.strptime(date, '%m/%d/%Y'))])
            # Total Nominal COD
            cod_amount = df_cnote.loc[(df_cnote['Manifest Bag No'] == datetime.strptime(date, '%m/%d/%Y')) &
                                                   (df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['COD flag'] == 'Y') & (df_cnote['Hawb Customer Name'] == customer), 'COD amount'].sum()
            cust_cod.append([ship_cod, cod_amount])
        whole_cod.append(cust_cod)
    cod_all.append([all_cod, all_cod, whole_cod[0],
                    whole_cod[1], whole_cod[2]])

   # -------------------- ! -------------------- ! -------------------- ! -------------------- ! -------------------- ! --------------------

    # Hitung data H+0 - H+X
    data_harian = []
    for hari in range(0, 17):

        # Hitung data tiap CUSTOMER & ZONA
        # ---------- ALL ----------
        data_all = []
        for zona in range(jumlah_zona):
            nama_zona = sheet[f'C{4 + zona}'].value

            # Total Connote
            all_total = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (
                datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona)])

            # UnRunsheet
            all_unrunsheet = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNRUNSHEET')])

            # Sukses
            all_sukses = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'DELIVERED')])

            # CR
            all_cr = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'CUSTOMER REQUEST')])

            # Undel
            all_undel = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNDELIVERY')])

            # Unstatus
            all_unstatus = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'OPEN')])

            # WH1
            all_wh1 = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'WH1')])

            # IRREGULARITY
            all_irreg = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'IRREGULARITY')])

            # RETURN
            all_return = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'RETURN')])

            data_all.append([all_total, all_unrunsheet, all_sukses, all_cr,
                             all_undel, all_unstatus, all_wh1, all_irreg, all_return])

        # ---------- COD ----------
        data_cod = []
        for zona in range(jumlah_zona):
            nama_zona = sheet[f'C{4 + zona}'].value

            # Total Connote
            cod_total = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (
                datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['COD flag'] == "Y")])

            # UnRunsheet
            cod_unrunsheet = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNRUNSHEET') & (df_cnote['COD flag'] == "Y")])

            # Sukses
            cod_sukses = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'DELIVERED') & (df_cnote['COD flag'] == "Y")])

            # CR
            cod_cr = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'CUSTOMER REQUEST') & (df_cnote['COD flag'] == "Y")])

            # Undel
            cod_undel = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNDELIVERY') & (df_cnote['COD flag'] == "Y")])

            # Unstatus
            cod_unstatus = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'OPEN') & (df_cnote['COD flag'] == "Y")])

            # WH1
            cod_wh1 = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'WH1') & (df_cnote['COD flag'] == "Y")])

            # IRREGULARITY
            cod_irreg = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'IRREGULARITY') & (df_cnote['COD flag'] == "Y")])

            # RETURN
            cod_return = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'RETURN') & (df_cnote['COD flag'] == "Y")])

            data_cod.append([cod_total, cod_unrunsheet, cod_sukses, cod_cr,
                             cod_undel, cod_unstatus, cod_wh1, cod_irreg, cod_return])

        # ---------- LAZADA, ORDIVO, TOKOPEDIA ----------
        whole_data = []
        for customer in nama_customer:
            data_cust = []
            for zona in range(jumlah_zona):
                nama_zona = sheet[f'C{4 + zona}'].value

                # Total Connote
                cust_total = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (
                    datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Hawb Customer Name'] == customer)])
                # UnRunsheet
                cust_unrunsheet = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                    date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNRUNSHEET') & (df_cnote['Hawb Customer Name'] == customer)])
                # Sukses
                cust_sukses = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                    date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'DELIVERED') & (df_cnote['Hawb Customer Name'] == customer)])
                # CR
                cust_cr = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                    date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'CUSTOMER REQUEST') & (df_cnote['Hawb Customer Name'] == customer)])
                # Undel
                cust_undel = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                    date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNDELIVERY') & (df_cnote['Hawb Customer Name'] == customer)])
                # Unstatus
                cust_unstatus = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                    date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'OPEN') & (df_cnote['Hawb Customer Name'] == customer)])
                # WH1
                cust_wh1 = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                    date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'WH1') & (df_cnote['Hawb Customer Name'] == customer)])
                # IRREGULARITY
                cust_irreg = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                    date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'IRREGULARITY') & (df_cnote['Hawb Customer Name'] == customer)])
                # RETURN
                cust_return = len(df_cnote[(df_cnote['Hawb Destination Name'] == nama_cabang[nowIndex]) & (df_cnote['Manifest Bag No'] == (datetime.strptime(
                    date, '%m/%d/%Y') - timedelta(days=hari)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'RETURN') & (df_cnote['Hawb Customer Name'] == customer)])
                data_cust.append([cust_total, cust_unrunsheet, cust_sukses, cust_cr,
                                  cust_undel, cust_unstatus, cust_wh1, cust_irreg, cust_return])
            whole_data.append(data_cust)

        # Add all counted data
        data_harian.append(
            [data_all, data_cod, whole_data[0], whole_data[1], whole_data[2]])

    return [cod_all, data_harian]


def counting_customer(file_data, date, is_grouped):
    zone_list = ['A', 'B', 'C', 'D']
    # Process raw data or load an already grouped data
    if is_grouped == 0:
        df_cnote = grouping_daily_monitor(
            file_data=file_data, save_grouping=False, saved_as='', tanggal=date)
    else:
        df_cnote = pd.read_excel(file_data)

    # Hitung data COD Shipment & Amount H+0
    cod_all = []
    # ---------- ALL & COD ----------
    all_cod = []
    for zona in range(0, 4):
        nama_zona = zone_list[zona]

        # Total Ship COD
        ship_cod = len(df_cnote[(df_cnote['Hawb PCS'] == nama_zona) & (
            df_cnote['COD flag'] == 'Y') & (df_cnote['Manifest Bag No'] == datetime.strptime(date, '%m/%d/%Y'))])

        # Total Nominal COD
        cod_amount = df_cnote.loc[(df_cnote['Manifest Bag No'] == datetime.strptime(
            date, '%m/%d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['COD flag'] == 'Y'), 'COD amount'].sum()

        # Failed COD Amount
        failed_cod_amount = df_cnote.loc[(df_cnote['Manifest Bag No'] == (datetime.strptime(
            date, '%m/%d/%Y') - timedelta(days=16)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['COD flag'] == 'Y') & (df_cnote['Status group'] == "RETURN"), 'COD amount'].sum()
        all_cod.append([ship_cod, cod_amount, failed_cod_amount])

    # ---------- LAZADA, ORDIVO, TOKOPEDIA ----------
    whole_cod = []
    for customer in nama_customer:
        cust_cod = []
        for zona in range(0, 4):
            nama_zona = zone_list[zona]

            # Total Ship COD
            ship_cod = len(df_cnote[(df_cnote['Hawb PCS'] == nama_zona) & (
                df_cnote['COD flag'] == 'Y') & (df_cnote['Hawb Customer Name'] == customer) & (df_cnote['Manifest Bag No'] == datetime.strptime(date, '%m/%d/%Y'))])

            # Total Nominal COD
            cod_amount = cod_amount = df_cnote.loc[(df_cnote['Manifest Bag No'] == datetime.strptime(date, '%m/%d/%Y')) & (
                df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['COD flag'] == 'Y') & (df_cnote['Hawb Customer Name'] == customer), 'COD amount'].sum()
            cust_cod.append([ship_cod, cod_amount])
        whole_cod.append(cust_cod)
    cod_all.append([whole_cod[2],
                    whole_cod[0], whole_cod[1], all_cod, all_cod])

    # -------------------- ! -------------------- ! -------------------- ! -------------------- ! -------------------- ! --------------------

    # Hitung data H+0 - H+7, H+10, H+15, H+X
    data_harian = []
    for hari in range(0, 11):
        hari_ref = hari
        if hari == 8:
            hari_ref += 2
        elif hari == 9 or hari == 10:
            hari_ref += 6

        # ---------- ALL SHIPMENT ----------
        data_all = []
        for zona in range(0, 4):
            nama_zona = zone_list[zona]

            # Total Connote
            all_total = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona)])

            # UnRunsheet
            all_unrunsheet = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNRUNSHEET')])

            # Sukses
            all_sukses = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'DELIVERED')])

            # CR
            all_cr = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime(
                '%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'CUSTOMER REQUEST')])

            # Undel
            all_undel = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNDELIVERY')])

            # Unstatus
            all_unstatus = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'OPEN')])

            # WH1
            all_wh1 = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'WH1')])

            # IRREGULARITY
            all_irreg = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'IRREGULARITY')])

            # RETURN
            all_return = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'RETURN')])

            # BREACH
            all_breach = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'BREACH')])

            data_all.append([all_total, all_unrunsheet, all_sukses, all_cr,
                            (all_undel + all_wh1), all_unstatus, all_return, all_irreg, all_breach])

        # ---------- COD ----------
        data_cod = []
        for zona in range(0, 4):
            nama_zona = zone_list[zona]

            # Total Connote
            cod_total = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(
                date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['COD flag'] == 'Y')])

            # UnRunsheet
            cod_unrunsheet = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNRUNSHEET') & (df_cnote['COD flag'] == 'Y')])

            # Sukses
            cod_sukses = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'DELIVERED') & (df_cnote['COD flag'] == 'Y')])

            # CR
            cod_cr = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime(
                '%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'CUSTOMER REQUEST') & (df_cnote['COD flag'] == 'Y')])

            # Undel
            cod_undel = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNDELIVERY') & (df_cnote['COD flag'] == 'Y')])

            # Unstatus
            cod_unstatus = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'OPEN') & (df_cnote['COD flag'] == 'Y')])

            # WH1
            cod_wh1 = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'WH1') & (df_cnote['COD flag'] == 'Y')])

            # IRREGULARITY
            cod_irreg = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'IRREGULARITY') & (df_cnote['COD flag'] == 'Y')])

            # RETURN
            cod_return = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'RETURN') & (df_cnote['COD flag'] == 'Y')])

            # BREACH
            cod_breach = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'BREACH') & (df_cnote['COD flag'] == 'Y')])

            data_cod.append([cod_total, cod_unrunsheet, cod_sukses, cod_cr,
                            (cod_undel + cod_wh1), cod_unstatus, cod_return, cod_irreg, cod_breach])

        # ---------- LAZADA, ORDIVO, TOKOPEDIA ----------
        whole_data = []
        for customer in nama_customer:
            data_cust = []
            for zona in range(0, 4):
                nama_zona = zone_list[zona]

                # Total Connote
                cust_total = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(
                    days=hari_ref)).strftime('%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Hawb Customer Name'] == customer)])

                # UnRunsheet
                cust_unrunsheet = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime(
                    '%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNRUNSHEET') & (df_cnote['Hawb Customer Name'] == customer)])

                # Sukses
                cust_sukses = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime('%#m/%#d/%Y')) & (
                    df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'DELIVERED') & (df_cnote['Hawb Customer Name'] == customer)])

                # CR
                cust_cr = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime('%#m/%#d/%Y')) & (
                    df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'CUSTOMER REQUEST') & (df_cnote['Hawb Customer Name'] == customer)])

                # Undel
                cust_undel = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime('%#m/%#d/%Y')) & (
                    df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'UNDELIVERY') & (df_cnote['Hawb Customer Name'] == customer)])

                # Unstatus
                cust_unstatus = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime(
                    '%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'OPEN') & (df_cnote['Hawb Customer Name'] == customer)])

                # WH1
                cust_wh1 = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime(
                    '%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'WH1') & (df_cnote['Hawb Customer Name'] == customer)])

                # IRREGULARITY
                cust_irreg = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime('%#m/%#d/%Y')) & (
                    df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'IRREGULARITY') & (df_cnote['Hawb Customer Name'] == customer)])

                # RETURN
                cust_return = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime(
                    '%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'RETURN') & (df_cnote['Hawb Customer Name'] == customer)])

                # BREACH
                cust_breach = len(df_cnote[(df_cnote['Manifest Bag No'] == (datetime.strptime(date, '%m/%d/%Y') - timedelta(days=hari_ref)).strftime(
                    '%#m/%#d/%Y')) & (df_cnote['Hawb PCS'] == nama_zona) & (df_cnote['Status group'] == 'BREACH') & (df_cnote['Hawb Customer Name'] == customer)])

                data_cust.append([cust_total, cust_unrunsheet, cust_sukses, cust_cr,
                                  (cust_undel + cust_wh1), cust_unstatus, cust_return, cust_irreg, cust_breach])
            whole_data.append(data_cust)
        data_harian.append(
            [whole_data[2], whole_data[0], whole_data[1], data_cod, data_all])

    return [cod_all, data_harian]
