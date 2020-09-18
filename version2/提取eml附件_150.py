import email
import os
from email.header import decode_header
import re

datapath = "c:/users/minu/desktop/minu109/"
names = os.listdir(datapath)


notannex = [] #不存在附件
notdecode = []  #名字不能解码
annex_sum = 0  #附件总数


#主过程，遍历所有eml
for index, i in enumerate(names):

    #补全路径
    filepath = os.path.join(datapath, i)      
    #print(filepath)


    #匹配poid
    poid = i[:14]

##    if len(poid) == 0:
##        print("......不存在poid")
##        break
##    poid = poid[0]
    
    #print(index)

    
    #解析邮件
    fp = open(filepath, "rb")
    #print(fp)
    msg = email.message_from_binary_file(fp)  #变成Email.message.message对象,以后补充


    annex = 0  #计算每个poid(邮件)的附件数
    
    #遍历mime的每一数据块
    for part in msg.walk():
        #multipart是指邮件分段边界标识,不在内容范围，不再执行下面解析 continue
        if part.get_content_maintype() == "multipart":
            
            #print("跳过无用块")
            continue

        #获取名称，根据名称保存附件
        annex_name = part.get_filename()  
        #print(annex_name)


        #附件存在名字，解码名字
        if annex_name:


            #解码，结果是list, 只有一个元素，元素为元组(value,charset)， 值 和 编码方式
            annex_name = decode_header(annex_name)
            value, charset = annex_name[0]

            #如果存在编码方式
            if charset:
                #微软将gb2312和gbk映射为gb18030,不改会报错
                if charset == "gb2312": 
                    value = value.decode("gbk")
                else:
                    value = value.decode(charset)
                #print(value)
            else:
                print("======不存在编码方式，以原码命名" + value)
                notdecode.append({ "poid": poid, "name": value })
                



            #名字格式化，添加poid
            final_name = poid + "-" + value[-35:]
            
            #判断是否是excel表格
            lastname = value.split(".")[-1]

            data = part.get_payload(decode = True)  #解码数据
            
            #保存附件
            if lastname == "xlsx" or lastname == "xls":
                
                with open(os.path.join("c:/users/minu/desktop/xlsx109/", final_name),"wb") as f:
                    f.write(data)

                #附件数
                annex += 1
                
            else:
                with open(os.path.join("c:/users/minu/desktop/other_annex", final_name), "wb") as f:
                    f.write(data)

                
            
    annex_sum += annex

    #输出记录
    print(str(index + 1) +  "， " + poid + "的excel数为： " + str(annex) + "， 当前总数： " + str(annex_sum))
    
    if annex == 0:
        print("!!!!!!邮件不存在附件，已记录" + poid)
        notannex.append(poid)



with open("c:/users/minu/desktop/log/no_annex.txt", "w") as f:
    f.write(str(notannex))
with open("c:/users/minu/desktop/log/can_not_decode.txt", "w") as f:
    f.write(str(notdecode))

    

