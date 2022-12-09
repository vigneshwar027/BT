# a = '''Current Status Expires
# I94 Expires
# I-797 Expires
# NIV Max Out Date
# I-129S Expires
# Petition Expiration Date PED
# EAD Expiration
# AP Expiration
# DS-2019 Expires
# Visa To Date
# Visa Expires Date
# I-551 Permanent Resident Card Expires
# Re-Entry Permit Expiration
# Visa Priority Date
# Management Info Job Start Date
# Date Retired'''

# a = a.split('\n')

# li =[]
# for i in a:
#     x =  "'{}':str".format(i)
#     li.append(x)

# zz = ','.join(li)
# print(zz)


import pandas as pd,numpy as np
from datetime import datetime

df_tab1 = pd.read_excel('Source Data\Open Process Data - Morningstar.xlsx')

# print(df['Current Status Expires'])
# print(df['NIV Max Out Date'])


date_col = ['Case No',"Petitioner","Beneficiary Name","Current Status","Current Status Expires","I-797 Expires","Entry Status","I94 Expires","NIV Max Out Date","Petition Expiration Date PED","EAD Expiration","HR Info - Department","Management Info Employee ID","Management Info Job Title","Process Case No","Case Opened","Process Type","Process Reference","Application Filed","Final Action Status","Final Action Date","Summary Case Disposition"]

for col in date_col:
    for index,row in df_tab1.iterrows():
        print(type(row['NIV Max Out Date']))
    break

# for column in date_col:
#     for index,row in  enumerate(df_tab1[column]):
#         # print(str(type(row)))
#         if 'tslibs' in str(type(row)):
#             # print('hii')
#             df_tab1[column] = pd.to_datetime(df_tab1[column], format='%Y-%m-%d',errors='coerce').dt.date
#             # print(column)
#             break
#         else:
#             try:
#                 df_tab1[column][index] = datetime.strftime(df_tab1[column][index],'%m/%d/%Y')
#             except:
#                 pass



# for column in date_col:
#     for i in df_tab1[column]:
#         print(i)
        
#     print('\n\n\n')



        # print(str(type(df_tab1['NIV Max Out Date'][x])))  

# Timestamp


# for column in date_col:
#     for index,row in df_tab1.iterrows():

#         # print(type(df_tab1[column][index]))
#         print((df_tab1[column][index]))
#         df_tab1[column][index] = pd.to_datetime(df_tab1[column][index]).dt.date
#         print((df_tab1[column][index]))
        
        # df[col][x] = datetime.strptime(df[col][x],'%Y-%m-%d %H:%M:%S')
        # print(df[col][x])

