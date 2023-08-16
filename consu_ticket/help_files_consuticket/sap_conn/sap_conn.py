import win32com.client
import sys
import time

def run_sap_gui(objSess,transaccion,variante,fecha,path,filename):

    try:
        #Execute SAP GUI script
        print("Running GUI script")
        time.sleep(1)
        objSess.findById("wnd[0]/tbar[0]/okcd").Text = transaccion
        objSess.findById("wnd[0]").sendVKey(0)
        objSess.findById("wnd[0]/tbar[1]/btn[17]").press()
        objSess.findById("wnd[1]/usr/txtV-LOW").Text = variante
        objSess.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
        objSess.findById("wnd[1]/usr/txtV-LOW").caretPosition = 9
        objSess.findById("wnd[1]/tbar[0]/btn[8]").press()
        objSess.findById("wnd[0]/usr/ctxtS_DATE-LOW").Text = fecha
        objSess.findById("wnd[0]/usr/ctxtS_DATE-HIGH").Text = fecha
        objSess.findById("wnd[0]/usr/ctxtS_DATE-HIGH").SetFocus
        objSess.findById("wnd[0]/usr/ctxtS_DATE-HIGH").caretPosition = 10
        objSess.findById("wnd[0]/tbar[1]/btn[8]").press()
        objSess.findById("wnd[0]/tbar[1]/btn[45]").press()
        objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[1]/usr/ctxtDY_PATH").Text = path
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = filename
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 20
        objSess.findById("wnd[1]/tbar[0]/btn[11]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[12]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[12]").press()
    
    except Exception as e:
        print(str(e))
     

def open_conn_sap():
    try:
        #open connection with sap
        print("Open Connection with SAP")
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        connection = application.Children(0)
        objSess = connection.Children(0)

        return objSess

    except Exception as e:
        print(str(e))
        