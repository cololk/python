#-*- coding:utf-8 -*-
#本程式可比對兩個件件號之間SPEC CODE的差異

import openpyxl

def trans(cgetsheet):
	lis=[]
	for i in range(2,cgetsheet.max_row+1):
		lis.append(cgetsheet.cell(row=i, column=1).value)
	return set(lis)
    
wb=openpyxl.load_workbook('C:\\demo\\python\\B版改C版差異_AH_AK.xlsx')
sheetname=wb.get_sheet_names()
for i in range(0,len(sheetname)):
    getsheet=wb.get_sheet_by_name(sheetname[i])
    if i==0:
    	s1=trans(getsheet)
    elif i==1:
    	s2=trans(getsheet)
    elif i==2:
    	s3=trans(getsheet)
    else:
    	s4=trans(getsheet)
print("第一組差異為:")
print(s1.symmetric_difference(s2))
print("第二組差異為:")
print(s3.symmetric_difference(s4))
