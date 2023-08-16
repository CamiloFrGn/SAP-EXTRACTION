import win32com.client

sap_gui_auto = win32com.client.GetObject("SAPGUI")
application = sap_gui_auto.GetScriptingEngine
connection = application.Children(0)
objSess = connection.Children(0)

objSess.findById("wnd[0]/tbar[0]/okcd").Text = "z102b_pump_service"
objSess.findById("wnd[0]").sendVKey(0)
objSess.findById("wnd[0]/tbar[1]/btn[6]").press()
objSess.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").PressToolbarButton("&FIND")
objSess.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = "ASIGNACIONES"
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
objSess.findById("wnd[0]/usr/subSA_0001:SAPMZCX_CSDSLSMX1014_PUMP:0001/ctxtS_EDATU-LOW").Text = "11.02.2023"
objSess.findById("wnd[0]/usr/subSA_0001:SAPMZCX_CSDSLSMX1014_PUMP:0001/ctxtS_EDATU-HIGH").Text = "11.02.2023"
objSess.findById("wnd[0]/usr/subSA_0001:SAPMZCX_CSDSLSMX1014_PUMP:0001/ctxtS_EDATU-HIGH").SetFocus()
objSess.findById("wnd[0]/usr/subSA_0001:SAPMZCX_CSDSLSMX1014_PUMP:0001/ctxtS_EDATU-HIGH").caretPosition = 10
objSess.findById("wnd[0]").sendVKey(0)
objSess.findById("wnd[0]/usr/cntlCONT_PUMP_SERVICE/shellcont/shell").pressToolbarContextButton("&MB_VIEW")
objSess.findById("wnd[0]/usr/cntlCONT_PUMP_SERVICE/shellcont/shell").selectContextMenuItem ("&PRINT_BACK_PREVIEW")
objSess.findById("wnd[0]/tbar[1]/btn[33]").press()
objSess.findById("wnd[1]/tbar[0]/btn[71]").press()
objSess.findById("wnd[2]/usr/chkSCAN_STRING-START").Selected = False
objSess.findById("wnd[2]/usr/txtRSYSF-STRING").Text = "BOMB-DESP"
objSess.findById("wnd[2]/usr/chkSCAN_STRING-START").SetFocus()
objSess.findById("wnd[2]/tbar[0]/btn[0]").press()
objSess.findById("wnd[3]/usr/lbl[2,2]").SetFocus()
objSess.findById("wnd[3]/usr/lbl[2,2]").caretPosition = 6
objSess.findById("wnd[3]").sendVKey(2)
objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
objSess.findById("wnd[0]/tbar[1]/btn[45]").press()
objSess.findById("wnd[1]/tbar[0]/btn[0]").press()
objSess.findById("wnd[1]/usr/ctxtDY_PATH").Text = r"C:\Users\jsdelgadoc\Documents\Proyectos-Cemex\sap_extraction\text_files"
objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "progscac.txt"
objSess.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
objSess.findById("wnd[1]/tbar[0]/btn[11]").press()
objSess.findById("wnd[0]/tbar[0]/btn[12]").press()
objSess.findById("wnd[0]/tbar[0]/btn[12]").press()