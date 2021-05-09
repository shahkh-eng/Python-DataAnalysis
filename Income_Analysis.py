"""
pip install pandas
pip install xlrd
pip install xlsxWriter
"""
import pandas as pd
path = "C:\\Users\\HP\\PycharmProjects\\pandas\\data\\DayBook-Main Sheet.xlsx"
# will open specific sheet in the main Excel file
df = pd.read_excel(path, '23rd March Onwards Inc')

# Al hman Gdn
rgplotsale = df.loc[(df['Main Head'].str.contains('Al hman|AL hman')) & (df['Excel Head'].str.contains('Plot Payment|Plot Resale'))]
rgplottransfer = df.loc[(df['Main Head'].str.contains('Al hman|AL hman')) & (df['Excel Head'].str.contains('Transfer|transfer'))]
rgplottoken = df.loc[(df['Main Head'].str.contains('Al hman|AL hman')) & (df['Excel Head'].str.contains('Token|token'))]
rgplotcomm = df.loc[(df['Main Head'].str.contains('Al hman|AL hman')) & (df['Excel Head'].str.contains('commis|Commis'))]

#Ahmad Block
adplotsale = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('Plot Payment|Plot Sale'))]
adplottransfer = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('Transfer|transfer'))]
adplottoken = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('Token|token'))]
adplotcomm = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('commis|Commis'))]

#Bank Withdrawn
bwith = df.loc[(df['Main Head'].str.contains('Bank|bank')) | (df['Excel Head'].str.contains('Bank|bank'))]
#Cash in Hand
cash = df.loc[(df['Main Head'].str.contains('Cash|cash')) | (df['Excel Head'].str.contains('Cash in |cash in '))]
#Zam Zam
zamzamplotsale = df.loc[(df['Main Head'].str.contains('Zam Zam|zam zam')) & (df['Excel Head'].str.contains('Plot Payment|Plot Sale'))]
zamzamplottransfer = df.loc[(df['Main Head'].str.contains('Zam Zam|zam zam')) & (df['Excel Head'].str.contains('Transfer|transfer'))]

# Pivot Table
incomebyhead = pd.pivot_table(df, index = ['Main Head', 'Excel Head'], values = 'Amount', aggfunc='sum', margins= True, margins_name='Total')
incomebyhead.to_excel('data/incomes/Summary Table.xlsx', sheet_name='Income Summary')

# For making multiple sheets inside one Excel File with the custom names
#writer = pd.ExcelWriter('data/Income_Breakup.xlsx', engine='xlsxwriter')
writer = pd.ExcelWriter('data/Incomes/Income_Breakup.xlsx')
rgplotsale.to_excel(writer, sheet_name='RG Plot Sale', index=False)
rgplottransfer.to_excel(writer, sheet_name='RG Plot Transfer', index=False)
rgplottoken.to_excel(writer, sheet_name='RG Plot Token', index=False)
rgplotcomm.to_excel(writer, sheet_name='RG Plot Commission', index=False)

adplotsale.to_excel(writer, sheet_name='AD Plot Sale', index=False)
adplottransfer.to_excel(writer, sheet_name='AD Plot Transfer', index=False)
adplottoken.to_excel(writer, sheet_name='AD Plot Token', index=False)
adplotcomm.to_excel(writer, sheet_name='AD Plot Commission', index=False)

bwith.to_excel(writer, sheet_name='Bank Income', index=False)
cash.to_excel(writer, sheet_name='Cash in Hand', index=False)

zamzamplotsale.to_excel(writer, sheet_name='Zam Zam Plot Sale', index=False)
zamzamplottransfer.to_excel(writer, sheet_name='Zam Zam Plot Transfer', index=False)

writer.save()