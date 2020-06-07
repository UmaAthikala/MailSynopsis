# -*- coding: utf-8 -*-
"""
Created on Sat Jun  6 22:11:53 2020

@author: 14704
"""

import win32com.client
import win32com
import os
import sys
import MySQLdb



outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts;


def emailleri_al(folder):
    msgnum=4
    messages = folder.Items
    a=len(messages)
    if a>0:
        for messageitr in messages:
            if messageitr.UnRead==True:
              
               sender = messageitr.SenderEmailAddress
               body = messageitr.body
               subject= messageitr.subject
               #if sender != "":
#                       print("\n*******************",file=f)
                   #print(sender, file=f)
                   #print(body, file=f)
               print(sender)
               msgnum=msgnum+1
               print(msgnum)
               dbconnect = MySQLdb.connect("localhost", "root", "12345", "mail")
               cursor = dbconnect.cursor() 
               query = 'insert into mails values (%s,%s,%s,%s)'
               cursor.execute(query,(msgnum,sender,subject,body))
               dbconnect.commit()
               
                       
                      

   
            try:
               messageitr.Save
               messageitr.Close(0)
            except:
                pass



for account in accounts:
    global inbox
    inbox = outlook.Folders(account.DeliveryStore.DisplayName)
#    print("****Account Name**********************************",file=f)
#    print(account.DisplayName,file=f)
    print(account.DisplayName)
    if account.DisplayName=='uma.forcoding@outlook.com':
        folders = inbox.Folders
    
        for folder in folders:
            if folder.name=='Inbox':
#                print("****Folder Name**********************************", file=f)
#                print(folder, file=f)
#                print("*************************************************", file=f)
                emailleri_al(folder)
                a = len(folder.folders)
                print(folder.name);
        
print("Finished Succesfully")