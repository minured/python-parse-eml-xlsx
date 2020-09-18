import email
import os
from email.header import decode_header
import re
import openpyxl
import re

datapath = "c:/users/minu/desktop/minu109/"
names = os.listdir(datapath)


wb = openpyxl.Workbook()
ws = wb.active


        

for index, i in enumerate(names):
    poid = i[:14]
    print(i)

    filepath = os.path.join(datapath, i)


    fp = open(filepath, "rb")
    msg = email.message_from_binary_file(fp)

    try:
        to = msg.get("TO")
        to_list = re.findall(r"<(.*?)>", to)
    except:
        print("=======错误" + poid)
        
    if len(to_list) > 0:
        to_s = ""
        for j in to_list:
            
            to_s = to_s + j + "、"

        print(to_s)


    ws.append([poid, to_s])

wb.save("c:/users/minu/desktop/to109.xlsx")

    
