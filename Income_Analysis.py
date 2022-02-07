"""
pip install pandas
pip install xlrd
pip install xlsxWriter
"""
import pandas as pd
path = "C:\\Users\\HP\\PycharmProjects\\pandas\\data\\DayBook-Main Sheet.xlsx"
# will open specific sheet in the main Excel file
df = pd.read_excel(path, '23rd March Onwards Inc')

rgplotsale = df.loc[(df['Main Head'].str.contains('Al hman|AL hman')) & (df['Excel Head'].str.contains('Plot Payment|Plot Resale'))]
adplotsale = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('Plot Payment|Plot Sale'))]

# Pivot Table
incomebyhead = pd.pivot_table(df, index = ['Main Head', 'Excel Head'], values = 'Amount', aggfunc='sum', margins= True, margins_name='Total')
incomebyhead.to_excel('data/incomes/Summary Table.xlsx', sheet_name='Income Summary')

# For making multiple sheets inside one Excel File with the custom names
#writer = pd.ExcelWriter('data/Income_Breakup.xlsx', engine='xlsxwriter')
writer = pd.ExcelWriter('data/Incomes/Income_Breakup.xlsx')
rgplotsale.to_excel(writer, sheet_name='RG Plot Sale', index=False)
adplotsale.to_excel(writer, sheet_name='AD Plot Sale', index=False)

writer.save()
