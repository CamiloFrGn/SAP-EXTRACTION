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
#sys.path.insert(0, r'C:\Users\jsdelgadoc\Documents\otros\sap_extraction\pumping_service_summary')
sys.path.insert(0, r'C:\Users\E-JFRANCOGON\Downloads\sap_extraction\pumping_service_summary')
from help_files_bombeo.sql.connect_sql_server import query_sql_crud, send_df_to_sql, query_sql_df
from help_files_bombeo.sap_conn.sap_conn import open_conn_sap, run_sap_gui
from help_files_bombeo.clean_data.clean_data import clean_data

#run program
try:
    def main_bombeo(objSess):
        print("INITIAL EXECUTION STARTING FOR PUMPING SERVICE")
        #start timer for runtime
        start = timeit.default_timer()
        #order of days to consult
        days_list = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]
        today = datetime.now()
        for i in days_list:           
            timestamp = datetime.now()         
            timestamp = timestamp.strftime("%Y-%m-%d %H:%M:%S") #timestamp format
            date_to_extract = today - timedelta(days=i) #subtract days according to list value in execution
            dataframe_filter = date_to_extract.strftime("%Y-%m-%d") #convert to format for dataframe filter
            print("RUNNING FOR DATE: "+dataframe_filter+" // Timestamp: "+str(timestamp))
            sap_date = date_to_extract.strftime("%d.%m.%Y") #convert to format for SAP filter
            transaccion = "z102b_pump_service"
            variante = "ASIGNACIONES"
            disposicion = "BOMB-DESP"
            #path = path = r"C:\Users\C-COOCC003\Desktop\DatosMacro"
            path = r"C:\Users\E-JFRANCOGON\Downloads\sap_extraction\text_files"
            filename = "conser.txt" 
            database_name = "AT49_Bombeo_z102b"
            #database_name = "bombeo_copy"   
            
            #run sap gui
            #open sap conn
            #objSess = open_conn_sap()
            #wb_path = r"C:\Users\jsdelgadoc\Documents\otros\condensado.xlsm"
            wb_path = r"C:\Users\E-JFRANCOGON\Downloads\sap_extraction\condensado.xlsm"
            run_sap_gui(objSess,transaccion,variante,sap_date,disposicion,path,filename,wb_path)

            #get nombre_cluster
            factories = query_sql_df( "select Centro, [Planta Unica] as NombrePlanta  from SCAC_AT1_NombreCluster where activo = 1", () )
            
            #start to clean data
            data = clean_data(factories,dataframe_filter,wb_path)
            if len(data) <= 1:
                print("Finalizando sin datos")
            else:
                #delete data from same date
                print("deleting date in sql")    
                query_sql_crud("delete from "+database_name+" where fechaentrega = ?", (dataframe_filter))

                #send dataframe to sql
                print("sending dataframe to sql")
                send_df_to_sql(data,database_name)

                print("Cargue realizado. Datos cargados: "+str(len(data)))

        #print info to cmd
        print("CICLO TOTAL FINALIZADO!!!!")
        time.sleep(5)
        stop = timeit.default_timer()
        runtime = ((stop - start)/60)
        print('Runtime: ', str(runtime)+" minutes") 
        print("-------------------------------")

   
except Exception as e:
    print(str(e))
    

