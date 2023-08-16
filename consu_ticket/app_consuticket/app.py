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
#sys.path.insert(0, r'C:\Users\jsdelgadoc\Documents\otros\sap_extraction\consu_ticket')
sys.path.insert(0, r'C:\Users\E-JFRANCOGON\Downloads\sap_extraction\consu_ticket')
from help_files_consuticket.sql.connect_sql_server import query_sql_crud, send_df_to_sql
from help_files_consuticket.sap_conn.sap_conn import open_conn_sap, run_sap_gui
from help_files_consuticket.clean_data.clean_data import clean_data

#run program
def main_consuticket(objSess):

    try:
        print("RUNNING PROGRAM FOR CONSU TICKET")
        #start timer for runtime
        start = timeit.default_timer()
    
        today = datetime.now()              
        timestamp = today.strftime("%Y-%m-%d %H:%M:%S") #timestamp format
        date_to_extract = today - timedelta(days=1) #get yestarday date
        #date_to_extract = today
        dataframe_filter = date_to_extract.strftime("%Y-%m-%d") #convert to format for dataframe filter
        print("RUNNING FOR DATE: "+dataframe_filter+" // Timestamp: "+str(timestamp))
        sap_date = date_to_extract.strftime("%d.%m.%Y") #convert to format for SAP filter
        transaccion = "Z1045_CONSU_TICKET2"
        variante = "OPT-RUTAS"
        
        #path = r"C:\Users\jsdelgadoc\Documents\otros\sap_extraction\text_files"
        path = r"C:\Users\E-JFRANCOGON\Downloads\sap_extraction\text_files"
        filename = "progscac.txt" 
        database_name = "AT51_Z1045_CONSU_TICKET2"
        #database_name = "consu_copy"   
        
        #run sap gui
        #open sap conn
        #objSess = open_conn_sap()
        run_sap_gui(objSess,transaccion,variante,sap_date,path,filename)

        #start to clean data
        data = clean_data(timestamp,dataframe_filter,path,filename)

        if len(data) <= 1:
            print("FINALIZADO SIN DATOS")
        else:
            #delete date
            query_sql_crud("delete from "+database_name+" where fechainicio = ?", (dataframe_filter))
            
            #send dataframe to sql
            print("sending dataframe to sql")
            send_df_to_sql(data,database_name)

        
        #print info to cmd
        #print("Cargue realizado. Datos cargados: "+str(len(data)))
        time.sleep(5)
        stop = timeit.default_timer()
        runtime = ((stop - start)/60)
        print('Runtime: ', str(runtime)+" minutes") 
        print("-------------------------------")

    except Exception as e:
        print(str(e))
   
#main_consuticket()   


    


