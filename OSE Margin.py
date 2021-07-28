#!/usr/bin/env python
# coding: utf-8

# In[2]:


import datetime
from tkinter import messagebox,Tk
import csv
import requests
from lxml import etree
import os
import win32com.client 
import re
Recipients = ['周禮強','張人友','李學晟','林暐翔','張誌修','陳鈺棻','黃寅榮','楊子輝','黃雅文'
              ,'蔡尚庭','藍立朋','楊欣曄','蘇雅鈴','王昭賢','許晉嘉','楊雅涵','鍾濰聲','陳伯維'] # 設定收件人,雅文偷偷把自己加進來了><
def send_mail(contract,margin1,margin2):    
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail_item = outlook.CreateItem(0) # 0: olMailItem
    
    for i in range(len(Recipients)):
        mail_item.Recipients.Add(Recipients[i])
    
    mail_item.Subject = f'{contract}保證金異動還沒改,上手:{margin1} 公司:{margin2}'

    mail_item.BodyFormat = 2          # 2: Html format
    mail_item.HTMLBody  = f'{contract}保證金異動還沒改'
    mail_item.Send()


root=Tk()
root.withdraw()
root.wm_attributes('-topmost',1)

Path = os.path.expanduser("~/Desktop/Pyfile/保證金檢核") #檔案位置
today = datetime.date.today()

f=open(f'{Path}/config.txt',encoding = 'utf-8')
file_Context=f.read().split()
f.close()
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNameSpace("MAPI")
try:
    account = namespace.Folders(file_Context[0])
    收件匣 = account.Folders[file_Context[1]]
    if len(file_Context) >2:
        for i,file in enumerate(file_Context):
            if i > 1:
                globals()[file] = eval(file_Context[i-1]).Folders[file]
except Exception as e:
    messagebox.showinfo(e,f"找不到{str(file_Context)}")

messages = eval(file_Context[-1]).Items
attachment =''
#下載公司保證金
for index,message in enumerate(messages):
    if '統一期貨(F008000)保證金轉檔'in message.Subject  and  message.Senton.date() == today:
        attachments = message.Attachments
        attachment = attachments.Item(1)
        for attachment in message.Attachments:
            attachment.SaveAsFile(os.path.join(Path, str(attachment)))
            
if attachment =='':
    messagebox.showinfo("提醒",'什麼都沒有QQ')
else:      
    #比對------------------------------------------------------------------
    #公司margin
    efile = [file for file in os.listdir() if 'F008000'in file]
    print(efile[-1])
    pfc = open(efile[-1],'rt' )
    pfcmargin=  csv.reader(pfc)
    for margin in pfcmargin:
        if 'JTI' in margin[1]:
            pfcjti =margin[7].strip()
        elif 'JNI' in margin[1]:
            pfcjni =margin[7].strip()
    pfc.close()

    #官網margin
    try:
        res = requests.get("https://www.jpx.co.jp/jscc/en/cash/futures/marginsystem/span_data.html")
        content = res.content.decode()
        html = etree.HTML(content)
        exchangejni = html.xpath('//body/div[1]/div[2]/div[2]/div/div/table[3]/tbody/tr[2]/td[3]/text()')
        exchangejti = html.xpath('/html/body/div[1]/div[2]/div[2]/div/div/table[3]/tbody/tr[3]/td[2]/text()')
        if  'yen'not in str(exchangejni) or 'yen' not in str(exchangejti):
            messagebox.showinfo("提醒",'抓到奇怪的資料了')

        if ''.join(re.findall(r'[0-9]',str(exchangejti))) == pfcjti.partition('.')[0] and     ''.join(re.findall(r'[0-9]',str(exchangejni))) == pfcjni.partition('.')[0]:
            messagebox.showinfo("提醒",'一切都正常: )')
        else:
            if ''.join(re.findall(r'[0-9]',str(exchangejti))) != pfcjti.partition('.')[0]: 
                send_mail('JTI',exchangejti,pfcjti)
            elif ''.join(re.findall(r'[0-9]',str(exchangejni))) != pfcjni.partition('.')[0]:
                send_mail('JNI',exchangejni,pfcjni)
        for file in efile:
            os.remove(file)
    except:
        messagebox.showinfo("提醒",'抓不到上手資料')




