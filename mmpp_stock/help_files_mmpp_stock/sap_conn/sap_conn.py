import win32com.client
import sys
import time
import clipboard

def run_sap_gui(objSess,transaccion,variante,path,filename,plantas_list_str):

    try:
        #Execute SAP GUI script
        print("Running GUI script")
        time.sleep(1)
        objSess.findById("wnd[0]").maximize()
        objSess.findById("wnd[0]/tbar[0]/okcd").text = transaccion
        objSess.findById("wnd[0]").sendVKey(0)
        objSess.findById("wnd[0]/tbar[1]/btn[17]").press()
        objSess.findById("wnd[1]/usr/txtV-LOW").text = variante
        objSess.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
        objSess.findById("wnd[1]/usr/txtV-LOW").caretPosition = 10
        objSess.findById("wnd[1]/tbar[0]/btn[8]").press()
        objSess.findById("wnd[0]").maximize()
        objSess.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press()
        objSess.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,0]").text = ""
        objSess.findById("wnd[1]/tbar[0]/btn[24]").press()
        objSess.findById("wnd[1]/tbar[0]/btn[8]").press()
        clipboard.copy(plantas_list_str)
        objSess.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press()
        objSess.findById("wnd[1]/tbar[0]/btn[24]").press()
        objSess.findById("wnd[1]/tbar[0]/btn[8]").press()
        objSess.findById("wnd[0]/tbar[1]/btn[8]").press()
        objSess.findById("wnd[0]/tbar[1]/btn[45]").press()
        objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
        objSess.findById("wnd[1]/tbar[0]/btn[11]").press()
        objSess.findById("wnd[0]").sendVKey(12)
        objSess.findById("wnd[0]/tbar[0]/btn[12]").press()


    except Exception as e:
        print(str(e))
        sys.exit()

def run_sap_gui2(objSess,transaccion,path,filename,mat_list_str):

    try:
        #Execute SAP GUI script
        print("Running GUI script2")
        time.sleep(1)
        objSess.findById("wnd[0]/tbar[0]/okcd").text = transaccion
        objSess.findById("wnd[0]").sendVKey(0)
        objSess.findById("wnd[0]").maximize()
        objSess.findById("wnd[0]/usr/ctxtMATNR-LOW").text = ""
        objSess.findById("wnd[0]").maximize()
        objSess.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press()
        objSess.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,0]").text = ""
        objSess.findById("wnd[1]/tbar[0]/btn[24]").press()
        objSess.findById("wnd[1]/tbar[0]/btn[8]").press()
        clipboard.copy(mat_list_str)
        objSess.findById("wnd[0]").maximize()
        objSess.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press()
        objSess.findById("wnd[1]/tbar[0]/btn[24]").press()
        objSess.findById("wnd[1]/tbar[0]/btn[8]").press()
        objSess.findById("wnd[0]/tbar[1]/btn[8]").press()
        objSess.findById("wnd[0]/tbar[1]/btn[43]").press()
        objSess.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
        objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 23
        objSess.findById("wnd[1]/tbar[0]/btn[11]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[12]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[12]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[12]").press()


    except Exception as e:
        print(str(e))
        sys.exit()


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
        sys.exit()