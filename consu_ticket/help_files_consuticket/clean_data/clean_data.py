import pandas as pd
import numpy as np
import sys
import os

def clean_data(now,now_date,path,filename):
    try:
        print("Cleaning data")
        #read conser.txt file to get data from sap
        #full_path = r"C:\Users\jsdelgadoc\Documents\otros\sap_extraction\text_files\progscac.txt"
        full_path = os.path.join(path,filename)
        data = pd.read_csv(full_path,sep='|', encoding='latin-1', on_bad_lines='skip', skiprows=3)

        #get total rows in dataframe to then filter out the first and last row that are non valid rows
        data_len = len(data.index)-4

        data = data.loc[1:data_len]

        #clean dataset
        data = data.reset_index(drop=True)

        #drop unnamed columns
        data = data.drop(columns=['Unnamed: 0'])
        data = data.drop(columns=['Unnamed: '+str(len(data.columns.values))])  
        
        #rename columns to sql columns
        data = data.rename(columns={
            data.columns[0]: 'FechaInicio', 
            data.columns[1]: 'FechaFin', 
            data.columns[2]: 'HoraInicio', 
            data.columns[3]: 'HoraFin', 
            data.columns[4]: 'OrdenProduccion', 
            data.columns[5]: 'Entrega',
            data.columns[6]: 'Centro', 
            data.columns[7]: 'TipoMaterial', 
            data.columns[8]: 'UnidadMedida',
            data.columns[9]: 'Estatus', 
            data.columns[10]: 'CantidadReal', 
            data.columns[11]: 'Material', 
            data.columns[12]: 'TextoBreveMaterial', 
            data.columns[13]: 'm3producidos', 
            data.columns[14]: 'DescTecnica'})

        if len(data) <= 1:
            print("FINALIZADO SIN DATOS")
        else:
    
            #covert datetime for date columns
            data['FechaInicio'] = data['FechaInicio'].astype(str)
            data['FechaInicio'] = data['FechaInicio'].str.replace('.','-')
            data['FechaInicio'] = data['FechaInicio'].replace(r'^\s*$', np.nan, regex=True)
            data['FechaInicio'] = data['FechaInicio'].astype(str)
            data['FechaInicio'] = data.FechaInicio.str.strip()
            data['FechaInicio'] = pd.to_datetime(data['FechaInicio'], format="%d-%m-%Y")

            data['FechaFin'] = data['FechaFin'].astype(str)
            data['FechaFin'] = data['FechaFin'].str.replace('.','-')
            data['FechaFin'] = data['FechaFin'].replace(r'^\s*$', np.nan, regex=True)
            data['FechaFin'] = data['FechaFin'].astype(str)
            data['FechaFin'] = data.FechaFin.str.strip()
            data['FechaFin'] = pd.to_datetime(data['FechaFin'], format="%d-%m-%Y")

            data['HoraInicio'] = data['HoraInicio'].astype(str)
            data['HoraInicio'] = data['HoraInicio'].str.replace('.','-')
            data['HoraInicio'] = data['HoraInicio'].replace(r'^\s*$', np.nan, regex=True)
            data['HoraInicio'] = data['HoraInicio'].astype(str)
            data['HoraInicio'] = data.HoraInicio.str.strip()
            data['HoraInicio'] = pd.to_datetime(data['HoraInicio'], format="%H:%M:%S")
            

            data['HoraFin'] = data['HoraFin'].astype(str)
            data['HoraFin'] = data['HoraFin'].str.replace('.','-')
            data['HoraFin'] = data['HoraFin'].replace(r'^\s*$', np.nan, regex=True)
            data['HoraFin'] = data['HoraFin'].astype(str)
            data['HoraFin'] = data.HoraFin.str.strip()
            data['HoraFin'] = pd.to_datetime(data['HoraFin'], format="%H:%M:%S")


            #replace commas for numeric values
            data['CantidadReal'] = data['CantidadReal'].astype(str)
            data['CantidadReal'] = data.CantidadReal.str.strip()
            data['CantidadReal'] = data['CantidadReal'].str.replace('.','')
            data['CantidadReal'] = data['CantidadReal'].str.replace(',','.')
            data['m3producidos'] = data['m3producidos'].astype(str)
            data['m3producidos'] = data.m3producidos.str.strip()
            data['m3producidos'] = data['m3producidos'].str.replace('.','')   
            data['m3producidos'] = data['m3producidos'].str.replace(',','.')   
        

            #filter out dates that are not current date
            data = data[(data['FechaInicio'] == now_date)]   

            #datatype setup
            data['OrdenProduccion'] = data['OrdenProduccion'].astype(str)
            data['Entrega'] = data['Entrega'].astype(str)
            data['Centro'] = data['Centro'].astype(str)
            data['TipoMaterial'] = data['TipoMaterial'].astype(str)
            data['UnidadMedida'] = data['UnidadMedida'].astype(str)
            data['Estatus'] = data['Estatus'].astype(str)
            data['CantidadReal'] = data['CantidadReal'].astype(float)
            data['Material'] = data['Material'].astype(str)
            data['TextoBreveMaterial'] = data['TextoBreveMaterial'].astype(str)
            data['m3producidos'] = data['m3producidos'].astype(float)
            data['DescTecnica'] = data['DescTecnica'].astype(str)
                  
            
            #strip string columns
            data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            #replace blanks for null values
            data = data.replace(r'^\s*$', np.nan, regex=True)
            
            data['CantidadReal'] = data['CantidadReal'].fillna(0)
            data['m3producidos'] = data['m3producidos'].fillna(0)
            
        
        return data
    except Exception as e:
        print(str(e))
       