import xlrd
import os
import shutil
import openpyxl

fp = "c:/users/minu/desktop/xlsx/"
fl = os.listdir(fp)

ftest = []
ftest.append(fl[0])
ftest.append(fl[1])


noid = [] #不存在任务id

merror = []

nwb = openpyxl.Workbook()
nsheet = nwb.active

head = ["poid", "任务id", "所属主体", "达人昵称"]
nsheet.append(head)


#遍历xlsx文件
for index, i in enumerate(fl):

    poid = i[:12]

    print("==========表格的进度：" + str(index))

    #打开workbook
    wb = xlrd.open_workbook(os.path.join(fp, i))
    names = wb.sheet_names()  #sheet的列表

    check_cell = True  #是否继续遍历单元格
    
    #遍历worksheet
    for j in range(len(names)):
        ws = wb.sheets()[j]
        print(".....sheet数量：" + str(len(names)) + "，当前：" + str(j+1))

        if check_cell:
            print("开始遍历单元格,文件： " + i + ", " + "表： " + ws.name)
            #遍历单元格
            col = ws.ncols
            row = ws.nrows
            print("行数：" + str(row) + "，列数： " + str(col))

            if row < 10:
                for r in range(row):

                    #检查是否已找到字段
                    if check_cell: 
                        print("=====每行====")
                        for c in range(col):
                            print("每列")

                            #任务id
                            if ws.cell_value(r,c) == "任务id" or ws.cell_value(r, c) == "任务ID" or ws.cell_value(r, c) == "任务 ID" or ws.cell_value(r, c) == "任务 id":
     
                                check_cell = False #标记停止遍历
                            
                                #找到任务id，确定行号，开始找遍历当前行的每列，找其他两个
                                
                                for minu_c in range(col):

                                    #重置主体列号
                                    zhuti_c = -1
                                    
                                    if ws.cell_value(r, minu_c) == "达人昵称":
                                        print(minu_c)
                                        nicheng_c = minu_c
                                    if ws.cell_value(r, minu_c) == "达人主体":
                                        zhuti_c = minu_c


                                        #三个全部确定完
                                        print("已确定行号：" + str(r) + "任务id列：" + str(c) + "，昵称列：" + str(nicheng_c) + "，主体列：" + str(zhuti_c))
                                        n = 1
                                        while r + n < row:  #下面行数不超过总行数

                                            #任务id
                                            renwuid = str(ws.cell_value(r + n, c)).split(".")[0]
                                            print(renwuid)

                                            #达人昵称
                                            nicheng = str(ws.cell_value(r + n, nicheng_c))
                                            zhuti = str(ws.cell_value(r + n, zhuti_c))

                                            #开始写入表格
                                            nsheet.append([poid, renwuid, zhuti, nicheng])
                                            n += 1
                                            
                                break  #进跳出当前行的，遍历列


                                if zhuti_c == -1:
                                    print("无法找到主体" + str(poid))
                                    merror.append(poid)
            else:
                merror.append(poid)
                            
nwb.save("c:/users/minu/desktop/fill_v1.xlsx")



with open("c:/users/minu/desktop/error_v1.txt", "w") as f:
    f.write(str(set(merror)))
    print("错误已保存 %d" %len(merror))
