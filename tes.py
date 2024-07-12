import pandas as pd


# ref = pd.read_excel(r'D:\INBOUND REF.xlsx', sheet_name='CUSTOMER')
data = pd.read_excel(r'D:\DATA INBOUND.xlsx')
# df = pd.DataFrame(ref)
df2 = pd.DataFrame(data)

# customer = {}

# for i in range(0, ref.shape[0]):
#     customer[str(ref['No. Account2']
#                  [i])] = ref['Cust Grouping'][i]

print(type(df2['Hawb Customer'][0]))
print(df2['Hawb Customer'][0])
