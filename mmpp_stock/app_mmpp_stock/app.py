#libraries
from datetime import datetime
import time
import timeit
import warnings
warnings.filterwarnings('ignore')
import sys
import clipboard
#our own py scripts
#insert path to get help files 
#sys.path.insert(0, r'C:\Users\jsdelgadoc\Documents\otros\mmpp_stock')
sys.path.insert(0, r'C:\Users\E-JFRANCOGON\Downloads\sap_extraction\mmpp_stock')
from help_files_mmpp_stock.sql.connect_sql_server import *
from help_files_mmpp_stock.sap_conn.sap_conn import *
from help_files_mmpp_stock.clean_data.clean_data import *
from app_mmpp_stock.proyeccion_stock_out import main_proyeccion

#run program
try:
    #open sap conn
    #objSess = open_conn_sap()
    def main_mmpp_stock(objSess):
        print("RUNNING PROGRAM FOR MMPP STOCK")
        #start timer for runtime
        start = timeit.default_timer()
                
        today = datetime.now()              
        timestamp = today.strftime("%Y-%m-%d %H:%M:%S") #timestamp format
        print("Timestamp: "+str(timestamp))
        transaccion = "MB52"
        transaccion2 = "MB5T"
        variante = "LOGBI_MMPP"
        path = r"C:\Users\E-JFRANCOGON\Downloads\sap_extraction\mmpp_stock\text_files_mmpp_stock"
        #path = r"C:\Users\jsdelgadoc\Documents\otros\mmpp_stock\text_files"
        filename = "mmpp_stock.txt" 
        filename2 = "mmpp_stock_transitos.txt" 
        database_name = "SCAC_AT52_MMPP_STOCK"
        database_name2 = "scac_at_mmpp_transitos"
        #database_name = "mmppstock_copy"  

        #get active plants
        get_plants_statement = "select distinct centro as planta  from SCAC_AT1_NombreCluster where pais = 'Colombia' and activo = 1"
        
        active_plants = query_sql_df(get_plants_statement,())
        active_plants = active_plants['planta'].unique()

        #get material from last month
        get_mat_statement = "select distinct material from AT51_Z1045_CONSU_TICKET2 where FechaInicio >= dateadd(year,-1,cast(getdate() as date)) and TipoMaterial like '%ADI%'"
    
        mat_list = query_sql_df(get_mat_statement,())
        mat_list = mat_list['material'].unique()
      
        
        #copy plant list to clipboard
        plantas_list_str = '\r\n'.join([str(x) for x in active_plants])
        mat_list_str = '\r\n'.join([str(x) for x in mat_list])
        clipboard.copy(mat_list_str)

        #run sap gui
        
        run_sap_gui(objSess,transaccion,variante,path,filename,plantas_list_str)
        time.sleep(3)
        clipboard.copy(plantas_list_str)
        run_sap_gui2(objSess,transaccion2,path,filename2,mat_list_str)

        #start to clean data
        data = clean_data(timestamp,path,filename)
        data2 = clean_data2(path,filename2)
     
        #delete data before insert for stock
        print("deleting data")
        delete_sql = "delete from "+database_name
        query_sql_crud(delete_sql,())

        #delete data before insert for stock
        print("deleting data2")
        delete_sql2 = "delete from "+database_name2
        query_sql_crud(delete_sql2,())

        #send dataframe to sql
        print("sending dataframe to sql")
        
        send_df_to_sql(data,database_name)

        #send dataframe2 to sql
        print("sending dataframe2 to sql")
        
        send_df_to_sql(data2,database_name2)

        #update material
        print("updating data")
        update_sql = "update "+database_name+" set material = replace(material,'.0','')"
        query_sql_crud(update_sql,())

        #update2 material
        print("updating data2")
        update_sql2 = "update "+database_name2+" set material = replace(material,'.0','')"
        query_sql_crud(update_sql2,())

        #update centros with double production line
        print("updating centros")
        procedure_name = 'scac_ap_actualizar_centros_doble_linea_mmpp'
        update_sql3 = "execute "+procedure_name
        query_sql_crud(update_sql3,())

        print("updating centros done")

        main_proyeccion()

        #print info to cmd
        
        time.sleep(2)
        stop = timeit.default_timer()
        runtime = ((stop - start)/60)
        print('Runtime: ', str(runtime)+" minutes") 
        print("-------------------------------")


     
 


    
except Exception as e:
    print(str(e))
    sys.exit()

