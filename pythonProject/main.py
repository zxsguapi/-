import time

import win32com.client as win32
import os

# 固定操作
excel=win32.Dispatch("Excel.Application")
excel.DisplayAlerts =False
excel.Visible =True
pwd=os.getcwd()

# 读取已经存在的excel表格
myexcel=excel.Workbooks.open(pwd+os.sep+"11月每周更新工作明细(当快钟鑫) .xls")
# 激活sheet
mysheet=myexcel.Worksheets("第一周")


# 输入要操作的id
# 循环次数=6
# while(循环次数>1):
#     str=input()
#     str2=str2+str
# list=str.split(",")
# #每次添加一个,把字符串遍历
# # id的数量

str2=""
cir=6
while(cir>0):
    str=input()
    str2=str2+str+','
    cir=cir-1

list=str2.split(",")
list=list[:len(list)-1]

idLen=len(list)

line=13


for id in range(idLen):
    mysheet.Cells(line,3).Value=list[id]
    mysheet.Cells(line,2).Value="钟鑫"
    line+=1


