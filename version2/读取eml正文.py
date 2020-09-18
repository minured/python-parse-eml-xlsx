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


print(ws)

#获取邮件的字符编码，首先在message中寻找编码，如果没有，就在header的Content-Type中寻找
def guess_charset(msg):
    charset = msg.get_charset()
    if charset == None:
        
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos+8:].strip()
    return charset

def get_content(msg,poid):
    for part in msg.walk():
        content_type = part.get_content_type()
        charset = guess_charset(part)

        
        #如果有附件，则直接跳过
        if part.get_filename()!=None:
            continue
        
        email_content_type = ''
        content = ''
        
        if content_type == 'text/plain':
            email_content_type = 'text'
        elif content_type == 'text/html':
            print('html 格式 跳过')
            continue #不要html格式的邮件
            email_content_type = 'html'
        if charset:
            try:
                content = part.get_payload(decode=True).decode(charset)
            #这里遇到了几种由广告等不满足需求的邮件遇到的错误，直接跳过了
            except AttributeError:
                print('type error')
            except LookupError:
                print("unknown encoding: utf-8")
        if email_content_type =='':
            continue
            #如果内容为空，也跳过
        #print(email_content_type + ' -----  ' + content)
        #邮件的正文内容就在content中

        #print(content)
        
        zhuti = re.findall("Dear(.*)", content)
        if len(zhuti) >0:
            print(zhuti[0])

            ws.append([poid, zhuti[0]])

    
        
for index, i in enumerate(names):
    poid = i[:14]
    print(i)

    filepath = os.path.join(datapath, i)


    fp = open(filepath, "rb")
    msg = email.message_from_binary_file(fp)
    #print(msg)
    
    #print(guess_charset(msg))
    try:
        get_content(msg, poid)
    except:
        print("===============错误")

wb.save("c:/users/minu/desktop/zhuti222.xlsx")
            
            

        
        
            
