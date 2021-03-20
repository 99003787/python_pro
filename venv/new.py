import pandas as pd
from openpyxl import load_workbook
#j=str(input("enter name"))
#k = str(input("enter email id"))
n=int(input("enter ps number"))
p=input("name: ")
q=input("gmail: ")
#sheet_name=['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4' , 'Sheet5']
z=pd.read_excel('pythonexcel2.xlsx', sheet_name=['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4' , 'Sheet5'])
y=z['Sheet1']
y=y[y['PS number'] == n]
y=y[y['Official Email Address'] == q]
y=y[y['Display Name'] == p]
df = pd.DataFrame(y, columns=['SL#', 'PS number', 'Display Name', 'Official Email Address'])
for i in z.keys():
    x=z[i]
    t = x[x['PS number']==n]
    t = x[x['Official Email Address'] == q]
    t = x[x['Display Name'] == p]
    col = x.columns
    for j in col:
        df[j]=t[j]
#df.to_excel('pythonexcel1.xlsx',sheet_name='master',index=False)
path = "pythonexcel2.xlsx"
book = load_workbook(path)
writer = pd.ExcelWriter(path, engine='openpyxl')
writer.book = book
# searching for the master sheet in all sheets
if 'Master Sheet' in book.sheetnames:
    ref = book['Master Sheet']
    # removing the previous data in the sheet
    book.remove(ref)
    # writing the data into the master sheet
df.to_excel(writer, sheet_name='Master Sheet')
# saving  the data in the master sheet
writer.save()
writer.close()

"""
enter ps number99003787
name: Arikatla Pujitha
gmail: arikatla.pujitha@ltts.com
"""