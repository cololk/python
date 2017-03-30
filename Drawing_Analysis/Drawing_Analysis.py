#-*- coding:utf-8 -*-
#本程式可進行圖譜的自動分析
import openpyxl

wb=openpyxl.load_workbook('AH_AK.xlsx')
sheetname=wb.get_sheet_names()
getsheet=wb.get_sheet_by_name(sheetname[0])

def GereratePartContent(num):
	PartContent={}
	for j in range(3, getsheet.max_row+1): # j是圖譜的列數
	        if getsheet.cell(row=j, column=getsheet.max_column-(2-num)).value==u'\u25cf':
	            for k in range(3, getsheet.max_column-3): # k是車規的欄數
	                if getsheet.cell(row=j, column=k).value==u'\u25cf':
	                    if getsheet.cell(row=j, column=1).value not in PartContent: # 如果此SPEC CODE是第一次在這個件號出現,就設定其值
	                        PartContent[getsheet.cell(row=j,column=1).value]=[getsheet.cell(row=2,column=k).value]
	                    else: #否則就改為添加其值
	                        PartContent[getsheet.cell(row=j,column=1).value].append(getsheet.cell(row=2,column=k).value)
	return  getsheet.cell(row=2, column=getsheet.max_column-(2-num)).value, PartContent

lis=[]
for i in range(3):
	lis.append(GereratePartContent(i)[0])
	lis.append(GereratePartContent(i)[1])
	print("件號序列 %s SPEC CODE組成內容為:%s" % (GereratePartContent(i)[0],GereratePartContent(i)[1]))

        
    
