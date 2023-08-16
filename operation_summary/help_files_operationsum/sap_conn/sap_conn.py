import win32com.client
import sys
import time
import xlwings as xw


def run_sap_gui(objSess,transaccion,variante,fecha,disposicion,path,filename,wb_path):

    try:
        #Execute SAP GUI script
        print("Running GUI script")
        time.sleep(1)
        objSess.findById("wnd[0]/tbar[0]/okcd").Text = transaccion
        objSess.findById("wnd[0]").sendVKey(0)
        objSess.findById("wnd[0]/tbar[1]/btn[17]").press()
        objSess.findById("wnd[1]/usr/txtV-LOW").Text = variante
        objSess.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
        objSess.findById("wnd[1]/usr/txtV-LOW").caretPosition = 10
        objSess.findById("wnd[1]/tbar[0]/btn[8]").press()
        objSess.findById("wnd[0]/usr/ctxtP_IDATE").Text = fecha
        objSess.findById("wnd[0]/usr/ctxtP_FDATE").Text = fecha
        objSess.findById("wnd[0]/usr/ctxtP_FDATE").SetFocus()
        objSess.findById("wnd[0]/usr/ctxtP_FDATE").caretPosition = 2
        objSess.findById("wnd[0]/tbar[1]/btn[8]").press()
        objSess.findById("wnd[0]/usr/cntlCC_CALL_OUT_REP/shellcont/shell").PressToolbarContextButton("&MB_VIEW")
        objSess.findById("wnd[0]/usr/cntlCC_CALL_OUT_REP/shellcont/shell").SelectContextMenuItem("&PRINT_BACK_PREVIEW")
        objSess.findById("wnd[0]/tbar[1]/btn[33]").press()
        objSess.findById("wnd[1]/tbar[0]/btn[71]").press()
        objSess.findById("wnd[2]/usr/chkSCAN_STRING-START").Selected = False
        objSess.findById("wnd[2]/usr/chkSCAN_STRING-RANGE").Selected = True
        objSess.findById("wnd[2]/usr/txtRSYSF-STRING").Text = disposicion
        objSess.findById("wnd[2]/usr/chkSCAN_STRING-RANGE").SetFocus()
        objSess.findById("wnd[2]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[3]/tbar[0]/btn[2]").press()
        objSess.findById("wnd[3]/usr/lbl[2,2]").SetFocus()
        objSess.findById("wnd[3]/usr/lbl[2,2]").caretPosition = 5
        objSess.findById("wnd[3]").sendVKey(2)
        objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[0]/tbar[1]/btn[45]").press()
        objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[1]/usr/ctxtDY_PATH").Text = path
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = filename
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
        objSess.findById("wnd[1]/tbar[0]/btn[11]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[15]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[15]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[15]").press()
        print("running excel macro")
        wb = xw.Book(wb_path)
        macro = wb.macro("Module1.StartExtract")
        macro()

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