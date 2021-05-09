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

# Al hman Gdn
rgenter = df.loc[(df['Main Head'].str.contains('Al hman|AL hman')) & (df['Excel Head'].str.contains('(?i)Entertainment'))]
rgsale = df.loc[(df['Main Head'].str.contains('Al hman|AL hman')) & (df['Excel Head'].str.contains('(?i)resale|commi|return|Payment|Purchase'))]
rgconstruction = df.loc[(df['Main Head'].str.contains('Al hman|AL hman')) & (df['Excel Head'].str.contains('(?i)Bridge|Severage|Sewerage|Diesel|Tractor|Repair|maintenance|boundry'))]
rgmisc = df.loc[(df['Main Head'].str.contains('Al hman|AL hman')) & (df['Excel Head'].str.contains('(?i)misc|combined|income|sadqa'))]

#Ahmad Block
adenter = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('(?i)Entertainment'))]
adldalegal = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('(?i)lda|legal'))]
adsale = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('(?i)resale|commi|return|Payment'))]
adnarowal = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('(?i)narowal'))]
adconstruction = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('(?i)Bridge|Severage|Sewerage|Diesel|Tractor|Repair|maintenance'))]
admisc = df.loc[(df['Main Head'].str.contains('Ahmad|ahmad')) & (df['Excel Head'].str.contains('(?i)misc|combined|income|sadqa'))]

#Land Purchase
landp = df.loc[(df['Main Head'].str.contains('(?i)Land|purchase'))]
#Personal
personal = df.loc[(df['Main Head'].str.contains('(?i)personal'))]
#Bank Deposit
bdeposit = df.loc[(df['Main Head'].str.contains('(?i)Bank'))]
#Cash in Hand
cash = df.loc[(df['Main Head'].str.contains('(?i)cash'))]
#Home
home = df.loc[(df['Main Head'].str.contains('(?i)home'))]

#Zam Zam
zamzammaint = df.loc[(df['Main Head'].str.contains('Zam Zam|zam zam')) & (df['Excel Head'].str.contains('(?i)Arif|Road|Nala|Office|Havely|Park|Plant|Repair|diesel|dera|repair|mainte'))]
zamzamenter = df.loc[(df['Main Head'].str.contains('Zam Zam|zam zam')) & (df['Excel Head'].str.contains('(?i)Entertainment|entertainment'))]
zamzamldalegal = df.loc[(df['Main Head'].str.contains('Zam Zam|zam zam')) & (df['Excel Head'].str.contains('(?i)Legal|LDA|NOC'))]
zamzamsale = df.loc[(df['Main Head'].str.contains('Zam Zam|zam zam')) & (df['Excel Head'].str.contains('(?i)return|Refund|Commi|Income|bill|misc'))]
zamzamsalaries = df.loc[(df['Main Head'].str.contains('Zam Zam|zam zam')) & (df['Excel Head'].str.contains('(?i)salar|Advance|sadqa'))]
zamzamikram = df.loc[(df['Main Head'].str.contains('Zam Zam|zam zam')) & (df['Excel Head'].str.contains('(?i)combined|ikram'))]

# Pivot Table
expbyhead = pd.pivot_table(df, index = ['Main Head', 'Excel Head'], values = 'Amount', aggfunc='sum', margins= True, margins_name='Total')
expbyhead.to_excel('data/expenses/Summary Table.xlsx', sheet_name='Expenses Summary')

# For making multiple sheets inside one Excel File with the custom names
writer = pd.ExcelWriter('data/expenses/Expenses_Breakup.xlsx')
rgenter.to_excel(writer, sheet_name='RG Entertainment', index=False)
rgsale.to_excel(writer, sheet_name='RG Sale Related', index=False)
rgconstruction.to_excel(writer, sheet_name='RG Construction', index=False)
rgmisc.to_excel(writer, sheet_name='RG Misc', index=False)

adenter.to_excel(writer, sheet_name='AD Entertainment', index=False)
adldalegal.to_excel(writer, sheet_name='AD LDA&Legal', index=False)
adsale.to_excel(writer, sheet_name='AD Sale', index=False)
adnarowal.to_excel(writer, sheet_name='Narowal', index=False)
adconstruction.to_excel(writer, sheet_name='Construction', index=False)
admisc.to_excel(writer, sheet_name='Misc', index=False)

landp.to_excel(writer, sheet_name='Land Purchase', index=False)
personal.to_excel(writer, sheet_name='Personal', index=False)
bdeposit.to_excel(writer, sheet_name='Bank Deposit', index=False)
cash.to_excel(writer, sheet_name='Cash in Hand', index=False)
home.to_excel(writer, sheet_name='Home', index=False)

zamzammaint.to_excel(writer, sheet_name='Zam Zam Maintenance', index=False)
zamzamenter.to_excel(writer, sheet_name='Zam Zam Entertainment', index=False)
zamzamldalegal.to_excel(writer, sheet_name='Zam Zam LDA&Legal', index=False)
zamzamsale.to_excel(writer, sheet_name='Zam Zam Sale', index=False)
zamzamsalaries.to_excel(writer, sheet_name='Zam Zam Salaries', index=False)
zamzamikram.to_excel(writer, sheet_name='Zam Zam Ikram Diary&Combined', index=False)

writer.save()