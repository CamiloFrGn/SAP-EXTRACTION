import sys
import os
import win32com.client as win32
import pandas as pd

def send_email(e):
    try:
        
        print("Sending email")
        #construct outlook app instance
        ol_app = win32.Dispatch('Outlook.Application')
        ol_ns = ol_app.GetNameSpace('MAPI')

        mail_item = ol_app.CreateItem(0)
        mail_item.Subject = "Error Script Programación"
        mail_item.BodyFormat = 1
        mail_item.Body = "Error!"
        mail_item.HTMLBody = "PRUEBA: Ha sucedido el siguiente error: "+str(e)
        mail_item.To = 'santiagoandres.ortiz@cemex.com; soniafernanda.manciper@cemex.com; juancamilo.francogonzalez@ext.cemex.com'
        mail_item.Save()
        mail_item.Send()
        
    except Exception as e:
        print(str(e))
        sys.exit()
