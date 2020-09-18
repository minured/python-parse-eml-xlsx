import re
import os
import openpyxl

path = "c:/users/minu/desktop/minu109"
names = os.listdir(path)


wb = openpyxl.Workbook()
nsheet = wb.active
print(nsheet)

head = ["PO ID", "任务ID", "所属主体", "达人呢称"]
nsheet.append(head)



##for index, i in enumerate(names):
##    poid = re.findall("po[0-9]{10}", i)
##    if len(poid) > 0:
##        print(index, poid[0])
##        nsheet.append(poid)
        


for index, i in enumerate(names):
    poids = []
    poid = i[:14]
    print(poid)
    poids.append(poid)
    nsheet.append(poids)
    
    
wb.save("poid150.xlsx")
print("%d行数据，保存完毕!" %(nsheet.max_row))   
        



