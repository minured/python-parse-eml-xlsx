import xlrd
import os
import shutil

fp = "c:/users/minu/desktop/xlsx109/"
fl = os.listdir(fp)


noid = [] #不存在任务id


#遍历xlsx文件
for index, i in enumerate(fl):

    key = False  #记录是否含有任务id
    print(index + 1)
    try:
        
        wb = xlrd.open_workbook(os.path.join(fp, i))
        names = wb.sheet_names()

        #遍历worksheet
        for j in range(len(names)):
            ws = wb.sheets()[j]

            #只要有任务id，停止遍历本sheet的单元格
            if key == False:
                 
                #遍历单元格
                col = ws.ncols
                row = ws.nrows
                for r in range(row):
                    for c in range(col):
                        if ws.cell_value(r,c) == "任务id" or ws.cell_value(r, c) == "任务ID" or ws.cell_value(r, c) == "任务 ID" or ws.cell_value(r, c) == "任务 id":
                            key = True
                            break
        if not key:
            print(i + "不存在任务id")
            
            shutil.move(os.path.join(fp, i), "c:/users/minu/desktop/norenwuid/")

    except:
        print("发生异常" + i)
        shutil.move(os.path.join(fp, i), "c:/users/minu/desktop/openerror/")

        
        

        


                
                    
                    

                
                
                                        
