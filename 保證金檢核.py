#!/usr/bin/env python
# coding: utf-8

# In[27]:


import requests
from lxml import etree
import os
import win32com.client as win32
desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
os.chdir(desktop) # 設定工作路徑為桌面
Recipients = ['周禮強','張人友','李學晟','林暐翔','張誌修','陳鈺棻','黃寅榮','楊子輝','黃雅文'
              ,'蔡尚庭','藍立朋','楊欣曄','蘇雅鈴','王昭賢','許晉嘉','楊雅涵','鍾濰聲','陳伯維'] # 設定收件人,雅文偷偷把自己加進來了><

def send_mail():
    
    outlook = win32.Dispatch('Outlook.Application')
    mail_item = outlook.CreateItem(0) # 0: olMailItem
    
    for i in range(len(Recipients)):
        mail_item.Recipients.Add(Recipients[i])
    
    mail_item.Subject = 'Nikkei225保證金異動'

    mail_item.BodyFormat = 2          # 2: Html format
    mail_item.HTMLBody  = '''
        <H2>https://www.jpx.co.jp/jscc/en/cash/futures/marginsystem/span_data.html</H2>
        OSE官網有異動，快去查看一下吧~ 
        '''
    mail_item.Send()
    
fp = open("Nikkei225.txt", "a")
res = requests.get("https://www.jpx.co.jp/jscc/en/cash/futures/marginsystem/span_data.html")
content = res.content.decode()
html = etree.HTML(content)
title = html.xpath('//body/div[1]/div[2]/div[2]/div/div/table[3]/tbody/tr[2]/td[3]/text()')
fp.write(str(title)) # 寫入margin
fp.close()

fpr = open("Nikkei225.txt", "r")
a = fpr.readline() 
a1 = a.split("]")[1] #取出第二次margin
a2 = a.split("]")[0] #取出第一次margin
fpr.close()

if a2 == a1:
    fpw = open("Nikkei225.txt", "w") 
    fpw.write(str(title))
    fpw.close()
else:
    if __name__ == '__main__':
        send_mail()
        fpw = open("Nikkei225.txt", "w") 
        fpw.write(str(title))
        fpw.close()      


# In[28]:


def send_mail1():
    
    outlook = win32.Dispatch('Outlook.Application')
    mail_item = outlook.CreateItem(0) # 0: olMailItem
    
    for i in range(len(Recipients)):
        mail_item.Recipients.Add(Recipients[i])
    
    mail_item.Subject = 'TOPIX保證金異動'

    mail_item.BodyFormat = 2          # 2: Html format
    mail_item.HTMLBody  = '''
        <H2>https://www.jpx.co.jp/jscc/en/cash/futures/marginsystem/span_data.html</H2>
        OSE官網有異動，快去查看一下吧~ 
        '''
    mail_item.Send()
    
fp = open("TOPIX.txt", "a")
res = requests.get("https://www.jpx.co.jp/jscc/en/cash/futures/marginsystem/span_data.html")
content = res.content.decode()
html = etree.HTML(content)
title = html.xpath('/html/body/div[1]/div[2]/div[2]/div/div/table[3]/tbody/tr[3]/td[2]/text()')
fp.write(str(title)) # 寫入margin
fp.close()

fpr = open("TOPIX.txt", "r")
a = fpr.readline() 
a1 = a.split("]")[1] #取出第二次margin
a2 = a.split("]")[0] #取出第一次margin
fpr.close()

if a2 == a1:
    fpw = open("TOPIX.txt", "w") 
    fpw.write(str(title))
    fpw.close()
else:
    if __name__ == '__main__':
        send_mail1()
        fpw = open("TOPIX.txt", "w") 
        fpw.write(str(title))
        fpw.close()      


# In[ ]:




