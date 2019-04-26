import pandas
import xlrd
import xlwt
from openpyxl import Workbook
from tempfile import TemporaryFile


col_num_vuids = 5
previous_voters = xlrd.open_workbook('Pct 3,5,10 (list)2019.xlsx',{'strings_to_numbers':True})
voters = xlrd.open_workbook('Voting History clean.xlsx')

book = xlwt.Workbook()
nonvoter_sheet = book.add_sheet('nonvoters')


sname = previous_voters.sheet_names()
vsname = voters.sheet_names()

previous_voters_sheet = previous_voters.sheet_by_name(sname[0])
voters_sheet = voters.sheet_by_name(vsname[0])
print(previous_voters_sheet.nrows)
print(voters_sheet.nrows)

vuids = []
for i in range(voters_sheet.nrows):
    vuids.append(voters_sheet.cell(i,col_num_vuids).value)
print(previous_voters_sheet.nrows)

nonvoters = []
for i in range(previous_voters_sheet.nrows):
    found = False
    for j in range(voters_sheet.nrows):
        if(previous_voters_sheet.cell(i,col_num_vuids-1).value == vuids[j]):
            found = True
    if(not found):
        print(previous_voters_sheet.getRow())
        #nonvoters.append(previous_voters_sheet.getRow(i))

print(nonvoters[0])
print(nonvoters[1])

for row,col in enumerate(nonvoters):
    #print(row[0])
    print(col)
    #nonvoter_sheet.write(row[0],nonvoters[row])

print(len(nonvoters))

name = 'nonvoters2019.xls'
book.save(name)
book.save(TemporaryFile())

'''
#riginal File
df = pandas.read_excel(r'./Pct 3,5,10 (list)2019.xlsx') #old 8692
df1 = pandas.read_excel(r'./Voting History clean.xlsx') #New 3850

df.columns = df.columns.str.strip().str.lower().str.replace(' ','_').str.replace('(','').str.replace(')','')
#df1.columns = df.columns.str.strip().str.lower().str.replace(' ','_').str.replace('(','').str.replace(')','')

df = df.query( "(precinct == 10)" )
df1 = df1.query("(Precinct == 10)" )
#df1.to_string()
vuids = df1["VUID"]
print(vuids[3841])



#for i in range(vuids.size):
    #if(df1.lookup([df1.index,vuids[i]])):
print(df.loc[df['vuid'].isin(vuids)])
#df.to_string()

#print(df)
#nonvoters = df.isin({},)

#print(df.columns)

#nonvoters = df - df1


#print(df[['vuid','precinct','name','status']])

#df.to_excel("nonvoters.xlsx")

#print(((df["Status"] != False) & (df["Precinct"] == 10)))
#print(df.query("(Status == 0) and (Precinct == 10)"))


'''