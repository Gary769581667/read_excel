
import numpy as np 
import xlrd
import xlwt

# timetable
workbook1 = xlrd.open_workbook("1_Timetable_Group_20(1).xlsx")
sheet1 = workbook1.sheet_by_index(0)
nrows1 = sheet1.nrows 
ncols1 = sheet1.ncols
A=[]
for i in range (nrows1-1):
    A.append(sheet1.row_values(i+1))
#duty period

workbook2 = xlrd.open_workbook("2.xlsx")
sheet2 = workbook2.sheet_by_index(0)
nrows2 = sheet2.nrows 
ncols2 = sheet2.ncols
B=[]
for i in range (nrows2-1):
    B.append(sheet2.row_values(i+1))

d={}
orig={}
dest={}
for i in range(nrows1-1):
    d[A[i][0]]=A[i][8]
    orig[A[i][0]]=A[i][1]
    dest[A[i][0]]=A[i][2]

wbk = xlwt.Workbook()
sheet = wbk.add_sheet('sheet 1')

for i in range(2020):
    s=B[i][0]
    s=s.split(',')

    sum = 0
    for j in range(len(s)):
        new_s = filter(str.isalnum, s[j])
        s[j]=''.join(list(new_s))
        sum = sum + d[s[j]]
    origin = orig[s[0]]
    destination = dest[s[-1]]

    sheet.write(i,0,sum)
    sheet.write(i,1,origin)
    sheet.write(i,2,destination)
# print(d['WA6171'],d['WA6197'],d['WA4308']+d['WA1107']+d['WA3232']+d['WA2262'],d['WA6074']+d['WA5388']+d['WA2429']+d['WA5089']) #172-176


wbk.save('total2.xls')





# matrix=np.zeros([nrows-1,ncols-1])

# for i in range(1,nrows):

#     data = sheet.row_values(rowx=i)
#     data.pop(0)
#     row_values = np.array(data)
#     matrix[i-1] = row_values


# print(matrix,"\n",matrix.shape)