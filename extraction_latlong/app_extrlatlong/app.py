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
#sys.path.insert(0, r'C:\Users\jsdelgadoc\Documents\otros\sap_extraction\extraction_latlong')
sys.path.insert(0, r'C:\Users\E-JFRANCOGON\Downloads\sap_extraction\extraction_latlong')
from help_files_extrlatlong.sql.connect_sql_server import query_sql_crud, send_df_to_sql
from help_files_extrlatlong.sap_conn.sap_conn import open_conn_sap, run_sap_gui
from help_files_extrlatlong.clean_data.clean_data import clean_data

#run program
def main_extrlatlong(objSess):

    try:
        print("RUNNING PROGRAM FOR EXTRACTION LAT LONG")
        #start timer for runtime
        start = timeit.default_timer()
    
        today = datetime.now()              
        timestamp = today.strftime("%Y-%m-%d %H:%M:%S") #timestamp format
        date_to_extract = today - timedelta(days=1) #get yestarday date
        #date_to_extract = today
        dataframe_filter = date_to_extract.strftime("%Y-%m-%d") #convert to format for dataframe filter
        print("RUNNING FOR DATE: "+dataframe_filter+" // Timestamp: "+str(timestamp))
        sap_date = date_to_extract.strftime("%d.%m.%Y") #convert to format for SAP filter
        transaccion = "Z1045_DELIVERY_SERV"
        variante = "LATLONGSCACV2"
        
        #path = r"C:\Users\jsdelgadoc\Documents\otros\sap_extraction\text_files"
        path = r"C:\Users\E-JFRANCOGON\Downloads\sap_extraction\text_files"
        filename = "progscac.txt" 
        database_name = "SCAC_AT8_MASTER_LAT_LONG"
        #database_name = "consu_copy"   
        
        #run sap gui
        #open sap conn
        #objSess = open_conn_sap()
        run_sap_gui(objSess,transaccion,variante,sap_date,path,filename)

        #start to clean data
        data = clean_data(path,filename)

        #send dataframe to sql
        #print("sending dataframe to sql")
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
      
 


    


