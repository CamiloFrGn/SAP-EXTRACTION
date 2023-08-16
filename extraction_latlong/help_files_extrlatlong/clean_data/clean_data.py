import pandas as pd
import numpy as np
import sys
import os

def clean_data(path,filename):
    try:
        print("Cleaning data")
        #read conser.txt file to get data from sap
        full_path = os.path.join(path, filename)
        #full_path = r"C:\Users\snortiz\Documents\projects\sap_extraction\text_files\progscac.txt"        
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
            data.columns[0]: 'Obra', 
            data.columns[1]: 'Latitud', 
            data.columns[2]: 'Longitud'})
        print(data)
        if len(data) <= 1:
            print("FINALIZADO SIN DATOS")
        else:
            

            data['Obra'] = data['Obra'].astype(str)
            data['Obra'] = data.Obra.str.strip()
            data['Latitud'] = data['Latitud'].astype(str)
            data['Latitud'] = data.Latitud.str.strip()
            data['Longitud'] = data['Longitud'].astype(str)
            data['Longitud'] = data.Longitud.str.strip()
        

          
            #datatype setup
            data['Latitud'] = data['Latitud'].astype(float)
            data['Longitud'] = data['Longitud'].astype(float)
            

            
            #strip string columns
            #data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            #replace blanks for null values
            data = data.replace(r'^\s*$', np.nan, regex=True)
            
            data['Latitud'] = data['Latitud'].fillna(0)
            data['Longitud'] = data['Longitud'].fillna(0)

            data['Latitud'] = data['Latitud'] * 100
            data['Longitud'] = data['Longitud'] * 100
            
    
        return data
    except Exception as e:
        print(str(e))
       