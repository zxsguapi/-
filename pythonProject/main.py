import time

import win32com.client as win32
import os

# 固定操作
excel = win32.Dispatch("Excel.Application")
excel.DisplayAlerts = False
excel.Visible = True
pwd = os.getcwd()

# 读取已经存在的excel表格
myexcel = excel.Workbooks.open(pwd + os.sep + "11月每周更新工作明细(当快钟鑫) .xls")
# 激活sheet
print("选择要操作的sheet")
sheet=input()
mysheet = myexcel.Worksheets(sheet)

# 输入要操作的id
# 循环次数=6
# while(循环次数>1):
#     str=input()
#     str2=str2+str
# list=str.split(",")
# #每次添加一个,把字符串遍历
# # id的数量

str2 = ""
str3 = ""

while True:
    print("输入要操作的id")
    str = input()
    print("输入execute开始执行excel，输入其他值继续插入数据")
    str3 = input()
    str2 = str2 + str + ','
    if (str3 == "execute"):
        break

list = str2.split(",")
list = list[:len(list) - 1]

idLen = len(list)

line = 13
print("输入要继续存还是重新存")
duandian=input()
while True:
    if(duandian=="继续"):
        if mysheet.Cells(line,3).Value is not None:
            line = line+1
            if mysheet.Cells(line,3).Value is None:
                break
    elif(duandian=="重新"):
        break

print(line)
for id in range(idLen):
    mysheet.Cells(line, 3).Value = list[id]
    mysheet.Cells(line, 2).Value = "钟鑫"
    line += 1

myexcel.SaveAs(pwd + os.sep+"11月每周更新工作明细(当快钟鑫) .xls")