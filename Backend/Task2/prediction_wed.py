import xlrd
import xlsxwriter
import openpyxl
from xlutils.copy import copy

workbook = xlrd.open_workbook('Data_wednesday.xlsx')
#wbk = openpyxl.load_workbook('Data_wednesday.xlsx')

sheet = workbook.sheet_by_index(0)
row_count = sheet.nrows
lis1=[]
dict1={}
dict2={}
time=1
time2=1

for i in range (1,row_count):
    sv = sheet.cell_value(i,2)
    lis1.append(sv)
for speed in lis1:
    dict2[time]=speed
    time+=1
for i in dict2:

    if dict2[i] < 3.3:
        dict2[i] = 0
    elif dict2[i] > 20:
        dict2[i] = 0
    else:
        dict2[i] =  0.5 * 1.23 * 0.4 * 8490 * dict2[i] * dict2[i] * dict2[i]/1000

rb = xlrd.open_workbook('Data_wednesday.xlsx')
wb = copy(rb)

w_sheet = wb.get_sheet(0)
ml = 1
for j in range(1,25):
    w_sheet.write(j,3,dict2[ml])
    wb.save('Data_wednesday.csv')
    ml += 1

# for i in dict2:
#     wbk.cell(row=2, column=4).value = dict2[i]
print(dict2)
