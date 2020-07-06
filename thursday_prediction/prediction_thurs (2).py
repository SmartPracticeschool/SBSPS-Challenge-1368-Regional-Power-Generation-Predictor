import xlrd
import xlsxwriter
import openpyxl
from xlutils.copy import copy

workbook = xlrd.open_workbook('Data_thursday.xlsx')
sheet = workbook.sheet_by_index(0)

row_count = sheet.nrows

lis1=[]
lis2=[]
lis3=[]
lis4=[]
dict1={}
dict2={}
dicttime={}
time=1
time2=1
rw_no=1
n1=1
n2=0

xyz=[]
def solar():
    n1 = -1
    n2 = -1
    n = -1 #index of lis3
    for i in range (1,row_count):
        solar5 = sheet.cell_value(i,1)#lis2=theory
        lis2.append(solar5)
    for i in range (1,row_count):
        solar3 = sheet.cell_value(i,0)
        lis3.append(solar3)#lis3=times

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
        dict2[i] =  (0.5 * 1.23 *2826 * dict2[i] * dict2[i] * dict2[i])/1000
rb = xlrd.open_workbook('Data_thursday.xlsx')
wb = copy(rb)
w_sheet = wb.get_sheet(0)
ml = 1
for j in range(1,25):
    w_sheet.write(j,4,dict2[ml])
    wb.save('Data_thursday.csv')
    ml += 1
need = 15000
sum = 0
for i in dict2:
    sum += dict2[i]
if sum <  need:
    print('There is shortfall of:',need - sum)
    solar()
    for i in lis2:
        if i == "Partly Cloudy":
            xyz.append(n1)
        n1+=1
elif sum > need:
    print('Excess Power Generated:', sum - need)
    n1 += 1
else:
    print('Requirements Stisfied:')
    n1 += 1
w_sheet.write(27,4,sum)
wb.save('Data_thursday.csv')
# print(lis2[n1])
print("Use solar at {} O'clock".format(xyz))

print(dict2)
