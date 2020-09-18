import xlrd
import os
import openpyxl


#获取文件列表
fptest = "xlsx109/"
desktop = "c:/users/minu/desktop/"
file_names = os.listdir(os.path.join(desktop, fptest))
#print(file_names)

my_error = []  #记录错误
multiple = []


#新建保存数据簿
save_wb = openpyxl.Workbook()
save_ws = save_wb.active
head = ["poid", "任务id", "达人主体", "达人昵称", "达人费用"]
save_ws.append(head)



#根据行和列获取值
def get_value(ws, row, renwuid_c, nicheng_c, zhuti_c, feiyong_c, save_ws, poid):
##    print("向下查找值")
    n = 1
    while row + n < ws.nrows:
##        print(row + n)

        #判断是否找到列号
        if renwuid_c != -1:
            renwuid = ws.cell_value(row + n, renwuid_c)
        else:
            renwuid = "没找到"
        if nicheng_c != -1:
            nicheng = ws.cell_value(row + n, nicheng_c)
        else:
            nicheng = "没找到"
        if zhuti_c != -1:
            zhuti = ws.cell_value(row + n, zhuti_c)
        else:
            zhuti = "没找到"
        if feiyong_c != -1:
            feiyong = ws.cell_value(row + n, feiyong_c)
        else:
            feiyong = "没找到"
            
##        print(renwuid, nicheng, zhuti)
        n += 1

        save_ws.append([poid, renwuid, zhuti, nicheng, feiyong])
        
        
    
#根据字段，确定字段所在行号
def get_main_row(ws, ws_i, field, fname):
    row_count = ws.nrows
    col_count = ws.ncols
    
##    print("[%d]：%s， 行数%d，列数%d" %(ws_i, ws.name, row_count, col_count))
    if row_count < 100:
        for r in range(row_count):
            for c in range(col_count):
                temp_value = ws.cell_value(r, c)   
                #print(r, c, temp_value)
                if temp_value == field or temp_value == "任务ID":
##                    print("找到任务id，确定行号：%d，%s-列号：%d" %(r, field, c))
                    return (r, c)
    else:
        multiple.append(fname + "-" + ws.name)



#查询其他字段的列号
def get_other_col(ws, row):
    nicheng_c = -1
    zhuti_c = -1
    feiyong_c = -1
##    print("    开始遍历当前行的列，确定其他字段")
    for col_i in range(ws.ncols):
        #print(row, col_i)
        temp_value = ws.cell_value(row, col_i)
        #print(temp_value)

        #达人昵称
        if temp_value == "达人昵称" or temp_value == "所属达人":
            #print("        达人昵称 %d,%d" %(row, col_i))
            nicheng_c = col_i
        elif temp_value == "达人主体" or temp_value == "所属机构":
            #print("        达人主体 %d,%d"%(row, col_i))
            zhuti_c = col_i
        elif temp_value == "达人费用" or temp_value == "费用":
            feiyong_c = col_i
            print(feiyong_c)

    return (nicheng_c, zhuti_c, feiyong_c)

               
#主流程，遍历文件
for index, fname in enumerate(file_names):
    poid = fname[:14]
    #打开工作簿，获取工作表的数量，遍历工作表
    wb = xlrd.open_workbook(os.path.join(desktop, fptest, fname))
    print("【 %d 】%s" %(index, fname))
    ws_names = wb.sheet_names()
    #print(ws_names)
    
 
    for ws_i in range(len(ws_names)):

        try:
            
            ws = wb.sheets()[ws_i]

            #查找 任务id 所在表头行和列
            row, renwuid_c = get_main_row(ws, ws_i, "任务id", fname)

            #根据表头行查找其他字段的行和列
            nicheng_c, zhuti_c, feiyong_c = get_other_col(ws, row)
            
            #print(nicheng_c, zhuti_c)

            #向下查找值
            get_value(ws, row, renwuid_c, nicheng_c, zhuti_c, feiyong_c, save_ws, poid)

        except:
            my_error.append(fname + "-" + ws.name)
            print("发生错误，已记录！")
            
    
    
save_wb.save(os.path.join(desktop, "fill_150_3333.xlsx"))
print("数据已保存！%d" %(save_ws.max_row))

with open(os.path.join(desktop,"error_v2.txt"), "w") as f:
    f.write(str(my_error))

with open(os.path.join(desktop,"more_than_100.txt"), "w") as f:
    f.write(str(multiple))
    print("错误已记录！error: %d，multiple: %d" %(len(my_error), len(multiple)))
