import pandas as pd
import numpy as np
import sys
import time

def clean_data(now,path,filename):
    try:
        print("Cleaning data")
        #read conser.txt file to get data from sap
        
        full_path = "/".join([path, filename])  

        data = pd.read_csv(full_path,sep='|', encoding='latin-1', skiprows=1)

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
            data.columns[0]: 'centro', 
            data.columns[1]: 'material', 
            data.columns[2]: 'descripcion_material', 
            data.columns[3]: 'unidad', 
            data.columns[4]: 'stock', 
            data.columns[5]: 'stock_transito', 
            data.columns[6]: 'stock_bloqueado'})

        if len(data) <= 1:
            print("FINALIZADO SIN DATOS")
        else:
             
            

            #replace commas for numeric values
            data['stock'] = data['stock'].astype(str)
            data['stock'] = data['stock'].str.replace('.','')
            data['stock'] = data['stock'].str.replace(',','.')
            data['stock_transito'] = data['stock_transito'].astype(str)
            data['stock_transito'] = data['stock_transito'].str.replace('.','')
            data['stock_transito'] = data['stock_transito'].str.replace(',','.')
            data['stock_bloqueado'] = data['stock_bloqueado'].astype(str)
            data['stock_bloqueado'] = data['stock_bloqueado'].str.replace('.','')
            data['stock_bloqueado'] = data['stock_bloqueado'].str.replace(',','.')
            
        
            #datatype setup
            print("CONVERTING COLUMNS DATA TYPE")
            data['centro'] = data['centro'].astype(str)
            data['material'] = data['material'].astype(str)

            data['descripcion_material'] = data['descripcion_material'].astype(str)
            data['unidad'] = data['unidad'].astype(str)

            #strip string columns
            data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)

            data['stock'] = data['stock'].astype(float)
            data['stock'] = data['stock'].fillna(0)
            data['stock_transito'] = data['stock_transito'].astype(float)
            data['stock_transito'] = data['stock_transito'].fillna(0)
            data['stock_bloqueado'] = data['stock_bloqueado'].astype(float)
            data['stock_bloqueado'] = data['stock_bloqueado'].fillna(0)


            #replace blanks for null values
            data = data.replace(r'^\s*$', np.nan, regex=True)

        return data
    except Exception as e:
        print(str(e))
        sys.exit()


def clean_data2(path,filename):
    try:
        print("Cleaning data2")
        #read conser.txt file to get data from sap
        
        full_path = "/".join([path, filename])  
        data = pd.read_csv(full_path,sep='|', encoding='latin-1', skiprows=3)

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
            data.columns[0]: 'centro_sumi', 
            data.columns[1]: 'material', 
            data.columns[2]: 'descripcion_material', 
            data.columns[3]: 'centro', 
            data.columns[4]: 'nombre_planta', 
            data.columns[5]: 'doc_compras', 
            data.columns[6]: 'pos',
            data.columns[7]: 'cantidad', 
            data.columns[8]: 'unidad'})

        if len(data) <= 1:
            print("FINALIZADO SIN DATOS")
        else:
         
            #replace commas for numeric values
            data['cantidad'] = data['cantidad'].astype(str)
            #data['cantidad'] = data['cantidad'].str.replace('.0','')
            #data['cantidad'] = data['cantidad'].str.replace(',','.')
            
    
            #datatype setup
            print("CONVERTING COLUMNS DATA TYPE")
            data['centro'] = data['centro'].astype(str)
            data['centro_sumi'] = data['centro_sumi'].astype(str)
            data['nombre_planta'] = data['nombre_planta'].astype(str)
            data['doc_compras'] = data['doc_compras'].astype(str)
            data['pos'] = data['pos'].astype(int)
            data['material'] = data['material'].astype(str)
            data['descripcion_material'] = data['descripcion_material'].astype(str)
            data['unidad'] = data['unidad'].astype(str)

            #strip string columns
            data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)

            
            data['cantidad'] = data['cantidad'].astype(float)
            data['cantidad'] = data['cantidad'].fillna(0)
            
            data.loc[data['cantidad'] < 10, 'cantidad'] = data['cantidad'] * 1000
            
            

            
            #replace blanks for null values
            data = data.replace(r'^\s*$', np.nan, regex=True)

        return data
    except Exception as e:
        print(str(e))
        sys.exit()
