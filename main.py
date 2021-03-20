import pandas as pd
Input = input("Enter the no of inputs you want:")
for i in range(int(Input)):
    h = int(input('Enter your ps no:-'))
    name = str(input('Enter the Name:-'))
    email = str(input('Enter the email:-'))
    z = pd.read_excel('pythonexcel2.xlsx', sheet_name=['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5'])
    y = z['Sheet1']
    y = y[(y['PS number'] == h) & (y['Display Name'] == name) & (y['Official Email Address'] == email)]
    if len(y) == 0:
        print('No match')
    else:
        df = pd.DataFrame(y, columns=['SL#', 'PS number', 'Display Name', 'Official Email Address'])
        for i in z.keys():
            x = z[i]
            t = x[(x['PS number'] == h) & (x['Display Name'] == name) & (x['Official Email Address'] == email)]
            col = x.columns
            for j in col:
                df[j] = t[j]
                df.to_excel('python_excel1.xlsx', sheet_name='master', index=False)