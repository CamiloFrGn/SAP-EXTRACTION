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
#sys.path.insert(0, r'C:\Users\jsdelgadoc\Documents\otros\sap_extraction')
sys.path.insert(0, r'C:\Users\E-JFRANCOGON\Downloads\sap_extraction')
from consu_ticket.app_consuticket.app import main_consuticket
from operation_summary.app_operationsum.app import main_operationsum
from pumping_service_summary.app_bombeo.app import main_bombeo
from mmpp_stock.app_mmpp_stock.app import main_mmpp_stock
from correct_adjustments.app_correct_adjustments.app import main_correct_adjustments
from extraction_latlong.app_extrlatlong.app import main_extrlatlong
#sys.path.insert(0, r'.\schedule_7_days')
sys.path.insert(0, r'C:\Users\E-JFRANCOGON\Downloads\sap_extraction\schedule_7_days')
from help_files_7days.sql.connect_sql_server import *
from help_files_7days.sap_conn.sap_conn import *
from help_files_7days.clean_data.clean_data import *
from help_files_7days.send_email.send_email import send_email
import pandas as pd
import socket
#run program
try:
    #open sap conn
    objSess = open_conn_sap()
    def main():
        try:
            print("RUNNING FOR 14 DAYS SCHEDULE")
            #start timer for runtime
            start = timeit.default_timer()
            #order of days to consult
            #days_list = [11,12,13,3,4,5,6,7,8,9,10,11,12,13,14,15,3,8,9,10,11,12,13,14,15,3,4,10,11,12,13,14,15,5,6,7,8,9,3,4,7,8,9,5,6] 
            days_list = [2,4,3,"5-15",0,4,1]
            today = datetime.now()
            for i in days_list:
                print("----------------------------------")
                hostname = socket.gethostname()
                query = "select * from scac_at_validacion_cargue where id_compu <> ?"

                validacion = query_sql_crud_select(query, (hostname)) 
                for x in validacion:
                    if str(x[1]) == str(i):
                        today = datetime.now()                   
                        dif = today - x[2]
                        dif = dif.total_seconds()/60

                        if dif < 10:
                            found = True
                            break
                        else:
                            found = False      
                    else:
                        found = False
                if found:
                    print("next i in list")
            
                else:
                    today = datetime.now()
                    query = "update scac_at_validacion_cargue set dia = ?, fecha_actualizacion = ? where id_compu = ?"

                    query_sql_crud_select(query, (i,today,hostname)) 
                 
                    timestamp = datetime.now()                  
                    #timestamp = timestamp.strftime("%Y-%m-%d %H:%M:%S") #timestamp format
                    if i == "5-15":
                        date_to_extract = today + timedelta(days=5) #add days according to list value in execution
                        date_to_extract2 = date_to_extract + timedelta(days=10)
                        dataframe_filter = date_to_extract.strftime("%Y-%m-%d") #convert to format for dataframe filter
                        dataframe_filter2 = date_to_extract2.strftime("%Y-%m-%d") #convert to format for dataframe filter
                        print("RUNNING FOR DATE: "+dataframe_filter+" to "+dataframe_filter2)
                        sap_date = date_to_extract.strftime("%d.%m.%Y")
                        sap_date1 = date_to_extract2.strftime("%d.%m.%Y") #convert to format for SAP filter
                
                    else:
                        date_to_extract = today + timedelta(days=i) #add days according to list value in execution
                        dataframe_filter = date_to_extract.strftime("%Y-%m-%d")
                        dataframe_filter2 = dataframe_filter #convert to format for dataframe filter
                        print("RUNNING FOR DATE: "+dataframe_filter)
                        sap_date = date_to_extract.strftime("%d.%m.%Y")
                        sap_date1 = sap_date #convert to format for SAP filter
                    
                    transaccion = "/nz102b_con_nser"
                    variante = "BASESCAC2"
                    transaccion2 = "Z102B_CALL_OUT"
                    variante2 = "ENLISTAESPERA"
                    disposicion = "PROGSCAC"
                    #path = r"C:\Users\jsdelgadoc\Documents\otros\sap_extraction\text_files"
                    path = r"C:\Users\E-JFRANCOGON\Downloads\sap_extraction\text_files"
                    filename = "conser.txt" 
                    database_name = "SCAC_AT15_Programacion7Dias"
                    #database_name = "programacion7dias_copy"
                    database_name2 = "scac_at_programacion_lista_espera"
                    #wb_path = r"C:\Users\jsdelgadoc\Documents\otros\programacion.xlsm"
                    wb_path = r"C:\Users\E-JFRANCOGON\Downloads\sap_extraction\condensado.xlsm"
                    
                    #run sap gui for waiting list
                    run_sap_gui2(objSess,transaccion2,variante2,sap_date,sap_date1,path,filename)
                    data2 = clean_data2(timestamp,dataframe_filter,dataframe_filter2,path,filename)
                    
                    #run sap guiprint
                    run_sap_gui(objSess,transaccion,variante,sap_date,sap_date1,disposicion,path,filename,wb_path)
                    time.sleep(2)
                    #start to clean data
                    timestamp = datetime.now()  
                    print("Timestamp: "+str(timestamp))  
                    data = clean_data(timestamp,dataframe_filter,dataframe_filter2,wb_path)
                    

                    #delete data from same date
                    
                    now = datetime.now()
                    now = now.minute         
                    now = now % 10
                    while now <= 2:
                        
                        time.sleep(10)
                        now = datetime.now()
                        now = now.minute         
                        now = now % 10
                        if now > 2:
                            break; 
                    time.sleep(1)
                    
                    print("deleting date in sql")
                    
                    if i == 0:
                        query_sql_crud("delete from "+database_name+" where fechaentrega <= ?", (dataframe_filter),data,database_name)
                        query_sql_crud("delete from "+database_name2+" where fechaentrega <= ?", (dataframe_filter),data2,database_name2)
                    else:
                        if i == "5-15":
                            query_sql_crud("delete from "+database_name+" where fechaentrega between ? and ?", (dataframe_filter,dataframe_filter2),data,database_name)
                            query_sql_crud("delete from "+database_name2+" where fechaentrega between ? and ?", (dataframe_filter,dataframe_filter2),data2,database_name2)
                        else:
                            query_sql_crud("delete from "+database_name+" where fechaentrega = ?", (dataframe_filter),data,database_name) 
                            query_sql_crud("delete from "+database_name2+" where fechaentrega = ?", (dataframe_filter),data2,database_name2)   
                    
                    
                    print("Cargue realizado. Datos cargados: "+str(len(data)))
                    time.sleep(5)
                
            #print info to cmd
            print("CICLO TOTAL FINALIZADO!!!!")
            time.sleep(5)
            stop = timeit.default_timer()
            runtime = ((stop - start)/60)
            print('Runtime: ', str(runtime)+" minutes") 
            print("-------------------------------")
        except Exception as e:
            print(str(e))
            send_email(e)
    def schedule_job():     
        now = datetime.now()
        now = now.strftime("%H:%M:%S")
        now = datetime.strptime(now, "%H:%M:%S").time()
        start_time_summary = "01:00:00"
        end_time_summary = "02:00:00"
        start_time_summary = datetime.strptime(start_time_summary, "%H:%M:%S").time()
        end_time_summary = datetime.strptime(end_time_summary, "%H:%M:%S").time()
        if now > start_time_summary and now < end_time_summary:
            print("#####################")
            main_consuticket(objSess)
            print("#####################")
            main_operationsum(objSess)
            print("#####################")
            main_bombeo(objSess)
            print("#####################")
            main_extrlatlong(objSess)
            print("#####################")
            main_correct_adjustments(objSess)
            print("#####################")
            print("sleeping...."+str(now))
            print("----------------------------------")
            main()
            print("#####################")
        start_time_mmpp_stock1 = "06:30:00"
        end_time_mmpp_stock1 = "07:15:00"
        start_time_mmpp_stock1 = datetime.strptime(start_time_mmpp_stock1, "%H:%M:%S").time()
        end_time_mmpp_stock1 = datetime.strptime(end_time_mmpp_stock1, "%H:%M:%S").time()
        start_time_mmpp_stock2 = "10:30:00"
        end_time_mmpp_stock2 = "11:15:00"
        start_time_mmpp_stock2 = datetime.strptime(start_time_mmpp_stock2, "%H:%M:%S").time()
        end_time_mmpp_stock2 = datetime.strptime(end_time_mmpp_stock2, "%H:%M:%S").time()
        start_time_mmpp_stock3 = "15:30:00"
        end_time_mmpp_stock3 = "16:15:00"
        start_time_mmpp_stock3 = datetime.strptime(start_time_mmpp_stock3, "%H:%M:%S").time()
        end_time_mmpp_stock3 = datetime.strptime(end_time_mmpp_stock3, "%H:%M:%S").time()
        if now > start_time_mmpp_stock1 and now < end_time_mmpp_stock1 or now > start_time_mmpp_stock2 and now < end_time_mmpp_stock2 or now > start_time_mmpp_stock3 and now < end_time_mmpp_stock3:
            print("#####################")
            main_mmpp_stock(objSess)
            print("#####################")
            main()

        else:        
            main()
    main()
    schedule.every(1).minutes.do(schedule_job)
    while True:
        schedule.run_pending()
        time.sleep(5)
    
except Exception as e:
    print(str(e))
    send_email(str(e))
    
    
    

