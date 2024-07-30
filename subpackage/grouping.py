import gc
import pandas as pd
from pathlib import Path
from datetime import datetime
from tkinter.messagebox import showinfo


def grouping_daily_monitor(file_data, tanggal, saved_as, save_grouping):
    # Load data source and reference table files
    df = pd.read_excel(file_data)
    working_dir = Path.cwd()
    df_zona = pd.read_excel(
        rf'{working_dir}\INBOUND REFS.xlsx', sheet_name='ZONA')
    df_destination = pd.read_excel(
        rf'{working_dir}\INBOUND REFS.xlsx', sheet_name='DESTINATION')
    df_customer = pd.read_excel(
        rf'{working_dir}\INBOUND REFS.xlsx', sheet_name='CUSTOMER')

    # Delete unused column
    df.drop(['Parameter Regional', 'Parameter Branch', 'Parameter Origin', 'Parameter Regional Dest.', 'Parameter Branch Dest.', 'Parameter Ring',
             'Parameter Destination', 'Parameter Date/Time', 'Parameter Date/Time From', 'Parameter Date/Time Thru', 'Jlc No', 'Goods Descr', 'Receiver Name',
             'Receiver Phone', 'Receiving Date', 'Hawb Branch Origin', 'Hawb Origin', 'Hawb Branch Destination', 'Hawb Amount', 'Hawb Packing', 'Hawb Cancel',
             'Hawb Type', 'Hawb Cust Type', 'Hawb Payment Type', 'Hawb Cust NA', 'Hawb Regional Dest.', 'Hawb Ring Dest.', 'Manifest UID', 'Zone', 'HVO No',
             'HVO Date', 'HVO Zone Dest', 'DO No', 'DO Date', 'RDO No', 'RDO Date', 'Pra Runsheet No', 'Pra Runsheet Date', 'Pra Runsheet Courier', 'DO',
            #  'First Runsheet Date', 'First Status', 'Second Runsheet Date', 'Second Status', 'Last Runsheet Date', 'Last Status'
             ],
            axis=1, inplace=True)

    try:
        # Process every row of data
        group = []
        sla = []
        zona_name = []
        customer_name = []
        destination_name = []
        inbound_date = []

        zona = {}
        customer = {}
        destination = {}

        # Load Reference Table for ZONA
        for i in range(0, df_zona.shape[0]):
            zona[str(df_zona['CODE'][i])] = df_zona['ZONA'][i]

        # Load Reference Table for DESTINATION
        for i in range(0, df_destination.shape[0]):
            destination[str(df_destination['CODE'][i])
                        ] = df_destination['NAMA KAB/KOTA'][i]

        # Load Reference Table for CUSTOMER
        for i in range(0, df_customer.shape[0]):
            customer[str(df_customer['No. Account2'][i])
                     ] = df_customer['Cust Grouping'][i]

        for index in range(0, df.shape[0]):
            # Status Group
            if df['Status wh'][index] == 'WH1' and str(df['POD Status'][index])[:1] == 'U':
                group.append('WH1')
            elif df['Status group'][index] == 'OTHERS' and df['POD Status'][index] == 'CR1':
                group.append('RETURN')
            elif df['Status group'][index] == 'OTHERS' and ((df['POD Status'][index] == 'D25' or df['POD Status'][index] == 'D26') or (df['POD Status'][index] == 'D37' or df['POD Status'][index] == 'R37') or df['POD Status'][index] == 'R25'):
                group.append('BREACH')
            elif df['Status group'][index] == 'OTHERS' and (df['POD Status'][index] == 'PS2' or (df['POD Status'][index] == 'PS3' or df['POD Status'][index] == 'UF')):
                group.append('IRREGULARITY')
            elif df['Status group'][index] == 'RETURN' and str(df['Runsheet Courier'][index])[:3] == "AMI":
                group.append('DELIVERED')
            elif df['Status group'][index] == "UNDELIVERY" and str(df['Runsheet Courier'][index])[:3] != "AMI":
                group.append("IRREGULARITY")
            elif df['Status group'][index] == "UNDEL - IRREGULARITY" or str(df['POD Status'][index])[:1] == 'U':
                group.append("UNDELIVERY")
            elif df['Status group'][index] == 'UNSTATUS':
                group.append("OPEN")
            elif pd.isna(df.at[index, 'Status group']):
                group.append("UNRUNSHEET")
            else:
                group.append(df['Status group'][index])

            # SLA
            tglh0 = datetime.strptime(
                tanggal,  "%m/%d/%Y").replace(hour=0, minute=0, second=0)
            tgl_data = (df['Manifest Inbound Date'][index]
                        ).replace(hour=0, minute=0, second=0)
            delta = tglh0 - tgl_data
            sla.append(delta.days)

            # HAWB Customer Name
            try:
                customer_name.append(customer[df['Hawb Customer'][index]])
            except:
                customer_name.append('#N/A')

            # HAWB Destination Name
            destination_name.append(
                destination[str(df['Hawb Destination'][index])])

            # Zona
            zona_name.append(zona[str(df['Hawb Destination'][index])])

            # Manifest Inbound Date
            int_date = (df['Manifest Inbound Date'][index]
                        ).replace(hour=0, minute=0, second=0)
            inbound_date.append(int_date)

        # Modify column value & append new column
        df.drop(['Status wh', 'Status group', 'Hawb Customer Name',
                'Hawb Destination Name', 'Hawb PCS', 'Manifest Bag No'], axis=1, inplace=True)
        df.insert(5, "Hawb Customer Name", customer_name)
        df.insert(8, "Hawb Destination Name", destination_name)
        df.insert(9, "Hawb PCS", zona_name)
        df.insert(15, "Manifest Bag No", inbound_date)
        df.insert(24, "Status group", group)
        df.loc[:, 'SLA'] = sla

        # Save the output or pass to other component
        if save_grouping == True:
            file_name = saved_as
            with pd.ExcelWriter(file_name, engine='xlsxwriter', date_format='m/d/yyyy') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)

            # Clean memory usage
            del df
            gc.collect()
            showinfo(title="Message",
                     message="Proses selesai")
        else:
            return df
    except Exception as e:
        showinfo(title="Error", message=f'{e}')
