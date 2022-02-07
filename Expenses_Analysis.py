"""
pip install pandas
pip install xlrd
pip install xlsxWriter
"""
# import pandas as pd
import pandas as pd
# excel file path
path = "C:\\Users\\HP\\PycharmProjects\\pandas\\data\\DayBook-Main Sheet.xlsx"
# will open specific sheet in the main Excel file
df = pd.read_excel(path, '23rd March Onwards Exp')

rgenter = df.loc[(df['Main Head'].str.contains('Al hman|AL hman')) & (df['Excel Head'].str.contains('(?i)Entertainment'))]
adenter = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('(?i)Entertainment'))]

#Land Purchase
landp = df.loc[(df['Main Head'].str.contains('(?i)Land|purchase'))]

zamzammaint = df.loc[(df['Main Head'].str.contains('Zam Zam|zam zam')) & (df['Excel Head'].str.contains('(?i)Arif|Road|Nala|Office|Havely|Park|Plant|Repair|diesel|dera|repair|mainte'))]

# Pivot Table
expbyhead = pd.pivot_table(df, index = ['Main Head', 'Excel Head'], values = 'Amount', aggfunc='sum', margins= True, margins_name='Total')
expbyhead.to_excel('data/expenses/Summary Table.xlsx', sheet_name='Expenses Summary')

# For making multiple sheets inside one Excel File with the custom names
writer = pd.ExcelWriter('data/expenses/Expenses_Breakup.xlsx')
rgenter.to_excel(writer, sheet_name='RG Entertainment', index=False)

writer.save()
