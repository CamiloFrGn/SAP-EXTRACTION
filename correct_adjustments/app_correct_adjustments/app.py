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
#sys.path.insert(0, r'C:\Users\jsdelgadoc\Documents\otros\sap_extraction\correct_adjustments')
sys.path.insert(0, r'C:\Users\E-JFRANCOGON\Downloads\sap_extraction\correct_adjustments')
from help_files_correct_adjustments.sql.connect_sql_server import query_sql_crud, send_df_to_sql
from help_files_correct_adjustments.sap_conn.sap_conn import open_conn_sap, run_sap_gui
from help_files_correct_adjustments.clean_data.clean_data import clean_data

#run program
def main_correct_adjustments(objSess):

    try:
        print("RUNNING PROGRAM FOR ADJUSTMENT CORRECTIONS")
        #start timer for runtime
        start = timeit.default_timer()
        variante_list = ["AJUSTESSCACV2","AJUSTESRDV2","AJUSTESPRV2","AJUSTESPANV2","AJUSTESNICV2","AJUSTESGTV2","AJUSTESCRV2"]
        for i in variante_list:
            today = datetime.now()              
            timestamp = today.strftime("%Y-%m-%d %H:%M:%S") #timestamp format
            date_to_extract = today - timedelta(days=1) #get yestarday date
            #date_to_extract = today
            dataframe_filter = date_to_extract.strftime("%Y-%m-%d") #convert to format for dataframe filter
            print(i+" RUNNING FOR DATE: "+dataframe_filter+" // Timestamp: "+str(timestamp))
            sap_date = date_to_extract.strftime("%d.%m.%Y") #convert to format for SAP filter
            transaccion = "Z1045_DELIVERY_SERV"
            variante = i
            
            #path = r"C:\Users\jsdelgadoc\Documents\otros\sap_extraction\text_files"
            path = r"C:\Users\E-JFRANCOGON\Downloads\sap_extraction\text_files"
            filename = "progscac.txt" 
            database_name = "SCAC_AT2_CondensadoServicio"
            #database_name = "consu_copy"   
            
            #run sap gui
            #open sap conn
            #objSess = open_conn_sap()
            run_sap_gui(objSess,transaccion,variante,sap_date,path,filename)

            #start to clean data
            data = clean_data(path,filename)
            if len(data) <= 1:
                print("no hay datos para actualizar")
            else:
                data = data.values.tolist()
              
                #update data
                print("updating data")
                query_sql_crud("update "+database_name+" set IDConductor = ?, NombreConductor = ?, Placa = ? where Entrega = ?;", data)
                

        #print info to cmd
        #print("Cargue realizado. Datos cargados: "+str(len(data)))
        time.sleep(5)
        stop = timeit.default_timer()
        runtime = ((stop - start)/60)
        print('Runtime: ', str(runtime)+" minutes") 
        print("-------------------------------")

    except Exception as e:
        print(str(e))
      



    


