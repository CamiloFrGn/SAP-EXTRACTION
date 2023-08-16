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
        objSess.findById("wnd[0]/tbar[1]/btn[6]").press()
        time.sleep(5)
        objSess.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").PressToolbarButton("&FIND")
        objSess.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = variante
        objSess.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").caretPosition = 13
        objSess.findById("wnd[2]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[2]/tbar[0]/btn[12]").press()
        objSess.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "28"
        objSess.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").DoubleClickCurrentCell()
        objSess.findById("wnd[0]").Maximize()
        objSess.findById("wnd[0]/usr/subSA_0002:SAPMZCX_CSDSLSMX1014_PUMP:0002/ctxtS_WERKS-LOW").Text = "F000"
        objSess.findById("wnd[0]/usr/subSA_0002:SAPMZCX_CSDSLSMX1014_PUMP:0002/ctxtS_WERKS-HIGH").Text = "G900"
        objSess.findById("wnd[0]/usr/subSA_0002:SAPMZCX_CSDSLSMX1014_PUMP:0002/ctxtS_WERKS-LOW").SetFocus()
        objSess.findById("wnd[0]/usr/subSA_0002:SAPMZCX_CSDSLSMX1014_PUMP:0002/ctxtS_WERKS-LOW").caretPosition = 3
        objSess.findById("wnd[0]/usr/subSA_0001:SAPMZCX_CSDSLSMX1014_PUMP:0001/ctxtS_EDATU-LOW").Text = fecha
        objSess.findById("wnd[0]/usr/subSA_0001:SAPMZCX_CSDSLSMX1014_PUMP:0001/ctxtS_EDATU-HIGH").Text = fecha
        objSess.findById("wnd[0]/usr/subSA_0001:SAPMZCX_CSDSLSMX1014_PUMP:0001/ctxtS_EDATU-HIGH").SetFocus()
        objSess.findById("wnd[0]/usr/subSA_0001:SAPMZCX_CSDSLSMX1014_PUMP:0001/ctxtS_EDATU-HIGH").caretPosition = 10
        objSess.findById("wnd[0]").sendVKey(0)
        objSess.findById("wnd[0]/usr/cntlCONT_PUMP_SERVICE/shellcont/shell").pressToolbarContextButton("&MB_VIEW")
        objSess.findById("wnd[0]/usr/cntlCONT_PUMP_SERVICE/shellcont/shell").selectContextMenuItem ("&PRINT_BACK_PREVIEW")
        objSess.findById("wnd[0]/tbar[1]/btn[33]").press()
        objSess.findById("wnd[1]/tbar[0]/btn[71]").press()
        objSess.findById("wnd[2]/usr/chkSCAN_STRING-START").Selected = False
        objSess.findById("wnd[2]/usr/txtRSYSF-STRING").Text = disposicion
        objSess.findById("wnd[2]/usr/chkSCAN_STRING-START").SetFocus()
        objSess.findById("wnd[2]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[3]/usr/lbl[2,2]").SetFocus()
        objSess.findById("wnd[3]/usr/lbl[2,2]").caretPosition = 6
        objSess.findById("wnd[3]").sendVKey(2)
        objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[0]/tbar[1]/btn[45]").press()
        objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[1]/usr/ctxtDY_PATH").Text = path
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = filename
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
        objSess.findById("wnd[1]/tbar[0]/btn[11]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[12]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[12]").press()
        print("running excel macro")
        wb = xw.Book(wb_path)
        macro = wb.macro("Module1.cargaInfoBombeo")
        macro()
    
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
        