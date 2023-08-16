#libraries
from datetime import datetime, timedelta
import time
import timeit
import warnings
warnings.filterwarnings('ignore')
import sys
import schedule
#our own py scripts
#insert path to get help files 
#sys.path.insert(0, r'C:\Users\jsdelgadoc\Documents\otros\sap_extraction\operation_summary')
sys.path.insert(0, r'C:\Users\E-JFRANCOGON\Downloads\sap_extraction\operation_summary')
from help_files_operationsum.sql.connect_sql_server import query_sql_crud, send_df_to_sql
from help_files_operationsum.sap_conn.sap_conn import open_conn_sap, run_sap_gui
from help_files_operationsum.clean_data.clean_data import clean_data


#run program
try:
    def main_operationsum(objSess):
        print("RUNNING PROGRAM FOR OPERATION SUMMARY")
        #start timer for runtime
        start = timeit.default_timer()
                
        today = datetime.now()              
        timestamp = today.strftime("%Y-%m-%d %H:%M:%S") #timestamp format
        date_to_extract = today - timedelta(days=1) #get yestarday date
        #date_to_extract = today
        dataframe_filter = date_to_extract.strftime("%Y-%m-%d") #convert to format for dataframe filter
        print("RUNNING FOR DATE: "+dataframe_filter+" // Timestamp: "+str(timestamp))
        sap_date = date_to_extract.strftime("%d.%m.%Y") #convert to format for SAP filter
        transaccion = "/nz102b_con_nser"
        variante = "BASESCAC2"
        disposicion = "PROGSCAC"
        path = r"C:\Users\E-JFRANCOGON\Downloads\sap_extraction\text_files"
        #path = r"C:\Users\C-COOCC003\Desktop\DatosMacro"
        
        filename = "conser.txt" 
        database_name = "SCAC_AT2_CondensadoServicio"
        #database_name = "condensado_full_copy"   
        
        #run sap gui
        #open sap conn
        #wb_path = r"C:\Users\jsdelgadoc\Documents\otros\condensado.xlsm"
        wb_path = r"C:\Users\E-JFRANCOGON\Downloads\sap_extraction\condensado.xlsm"
        #objSess = open_conn_sap()
        run_sap_gui(objSess,transaccion,variante,sap_date,disposicion,path,filename,wb_path)

        
        #start to clean data
        data = clean_data(timestamp,dataframe_filter,wb_path)
        
        #delete date
        query_sql_crud("delete from "+database_name+" where fechaentrega = ?", (dataframe_filter))

        #send dataframe to sql
        print("sending dataframe to sql")
        send_df_to_sql(data,database_name)

        #actualizar componentes de ciclo
        print("updating cycle components")
        procedure_name = 'scac_ap_actualizar_componentes_ciclo'
        query_sql_crud( 
                    "{CALL " + procedure_name+ " ()}", 
                    ( 
                    ) 
                )
        
        #print info to cmd
        print("Cargue realizado. Datos cargados: "+str(len(data)))
        time.sleep(5)
        stop = timeit.default_timer()
        runtime = ((stop - start)/60)
        print('Runtime: ', str(runtime)+" minutes") 
        print("-------------------------------")

    

    #schedule.every().day.at("02:45").do(main)
    #while True:
    #    schedule.run_pending()
    #    time.sleep(5)
    #main()
except Exception as e:
    print(str(e))
    sys.exit()

