import pandas as pd

customer = {}
customer_name = []

df = pd.read_excel(
    r"C:\Users\Lenovo\Documents\TOTAL TARIKAN INBOUND TANGGAL 17 - 02 JANUARI 2025.xlsx", nrows=100)

df_customer = pd.read_excel(
    r"C:\ERWIN\App Library\Python App\Daily Monitor Inbound\INBOUND REFS.xlsx", sheet_name='CUSTOMER')

# Load Reference Table for CUSTOMER
for i in range(0, df_customer.shape[0]):
    customer[str(df_customer['No. Account2'][i])
             ] = df_customer['Cust Grouping'][i]

for index in range(0, 10):
    # HAWB Customer Name
    try:
        customer_name.append(customer[str(df['Hawb Customer'][index])])
    except:
        customer_name.append('#N/A')

print(customer_name)
