import re
import os
import openpyxl

path = "c:/users/minu/desktop/minudata"
names = os.listdir(path)


wb = openpyxl.Workbook()
nsheet = wb.active

head = ["PO ID", "任务ID", "所属主体", "达人呢称"]
nsheet.append(head)



for index, i in enumerate(names):
    poid = re.findall("po[0-9]{10}", i)
    if len(poid) > 0:
        print(index, poid[0])
        nsheet.append(poid)
        


wb.save("poid800.xlsx")
print("%d行数据，保存完毕!" %(nsheet.max_row))   
        



