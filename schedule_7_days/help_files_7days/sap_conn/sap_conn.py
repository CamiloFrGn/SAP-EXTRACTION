import win32com.client
import sys
import time
import xlwings as xw
from help_files_7days.send_email.send_email import send_email

def run_sap_gui(objSess,transaccion,variante,fecha,fecha1,disposicion,path,filename,wb_path):

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
        objSess.findById("wnd[0]/usr/ctxtP_FDATE").Text = fecha1
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
        macro = wb.macro("Module1.StartExtract_programacion")
        macro()

    except Exception as e:
        print(str(e))
        send_email(str(e))
        sys.exit()

def run_sap_gui2(objSess,transaccion2,variante2,fecha,fecha1,path,filename):

    try:
        #Execute SAP GUI script
        print("Running GUI 2 script")
        time.sleep(1)
        objSess.findById("wnd[0]").maximize()
        objSess.findById("wnd[0]/tbar[0]/okcd").text = transaccion2
        objSess.findById("wnd[0]").sendVKey(0)
        objSess.findById("wnd[0]/tbar[1]/btn[17]").press()
        objSess.findById("wnd[1]/usr/txtV-LOW").text = variante2
        objSess.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        objSess.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
        objSess.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
        objSess.findById("wnd[1]/tbar[0]/btn[8]").press()
        objSess.findById("wnd[0]/usr/ctxtP_IDATE").text = fecha
        objSess.findById("wnd[0]/usr/ctxtP_FDATE").text = fecha1
        objSess.findById("wnd[0]/usr/ctxtP_FDATE").setFocus()
        objSess.findById("wnd[0]/usr/ctxtP_FDATE").caretPosition = 10
        objSess.findById("wnd[0]/tbar[1]/btn[8]").press()
        objSess.findById("wnd[0]/usr/cntlCC_CALL_OUT_REP/shellcont/shell").pressToolbarButton("&COL0")
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 83
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").firstVisibleRow = 74
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "0-83"
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_FL_SING").press()
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 55
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 48
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "55"
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 2
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 0
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "2"
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 65
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 60
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "65"
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 14
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 9
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "14"
        objSess.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
        objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[0]/usr/cntlCC_CALL_OUT_REP/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        objSess.findById("wnd[0]/usr/cntlCC_CALL_OUT_REP/shellcont/shell").selectContextMenuItem("&PC")
        objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
        objSess.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
        objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
        objSess.findById("wnd[1]/tbar[0]/btn[11]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[15]").press()
        objSess.findById("wnd[0]/tbar[0]/btn[15]").press()


    except Exception as e:
        print(str(e))
        send_email(str(e))
        sys.exit()

def run_sap_gui_error(objSess):

    try:
        #Execute SAP GUI script
        print("Running GUI script for error")
        time.sleep(1)
        objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
        

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