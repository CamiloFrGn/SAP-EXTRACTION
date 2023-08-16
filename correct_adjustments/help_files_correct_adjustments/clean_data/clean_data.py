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
        data_len = len(data.index)-2
      
        data = data.loc[1:data_len]
        #clean dataset
        data = data.reset_index(drop=True)
        #drop unnamed columns
        data = data.drop(columns=['Unnamed: 0'])
        data = data.drop(columns=['Unnamed: '+str(len(data.columns.values))])  
        
        #rename columns to sql columns
        data = data.rename(columns={
            data.columns[0]: 'entrega', 
            data.columns[1]: 'conductor', 
            data.columns[2]: 'nombre', 
            data.columns[3]: 'placa', 
            data.columns[4]: 'pos'})

        if len(data) <= 1:
            print("FINALIZADO SIN DATOS")
        else:

            print(data)
            data = data.drop(columns=['pos'])
            # Define the desired column order
            column_order = ['conductor', 'nombre', 'placa', 'entrega']

            # Reorder the columns
            data = data.reindex(columns=column_order)
         
            print(data)
            data['entrega'] = data['entrega'].astype(int)
            data['entrega'] = data['entrega'].astype(str)
            data['entrega'] = data.entrega.str.strip()
            data['conductor'] = data['conductor'].astype(int)
            data['conductor'] = data['conductor'].astype(str)
            data['conductor'] = data.conductor.str.strip()
            data['nombre'] = data['nombre'].astype(str)
            data['nombre'] = data.nombre.str.strip()
            data['placa'] = data['placa'].astype(str)
            data['placa'] = data.placa.str.strip()

            #strip string columns
            data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            #replace string blanks for null values
            data = data.replace(r'^\s*$', np.nan, regex=True)
            
            
            print(data)
        return data
    except Exception as e:
        print(str(e))
  