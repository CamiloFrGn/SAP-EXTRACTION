import pandas as pd
import numpy as np
import sys

def clean_data(factories,now_date,wb_path):
    try:
        print("Cleaning data")
        
        data = pd.read_excel(wb_path,sheet_name="Bombeo")
       
        #rename columns to sql columns

        if len(data) <= 1:
            print("FINALIZADO SIN DATOS")
        else:
            data = data.rename(columns={
                data.columns[0]: 'Centro', 
                data.columns[1]: 'FechaEntrega', 
                data.columns[2]: 'Volumen', 
                data.columns[3]: 'Pedido', 
                data.columns[4]: 'Posicion', 
                data.columns[5]: 'Equipo',
                data.columns[6]: 'Placa', 
                data.columns[7]: 'Entrega', 
                data.columns[8]: 'Cliente',
                data.columns[9]: 'Obra', 
                data.columns[10]: 'Denominacion', 
                data.columns[11]: 'Estatus'})
                    
            #covert datetime for date columns
            
            data['FechaEntrega'] = data['FechaEntrega'].astype(str)    
            data['FechaEntrega'] = data['FechaEntrega'].str.replace('.','-')
            data['FechaEntrega'] = data['FechaEntrega'].replace(r'^\s*$', np.nan, regex=True)
            data['FechaEntrega'] = data['FechaEntrega'].astype(str) 
            data['FechaEntrega'] = data.FechaEntrega.str.strip()
            data['FechaEntrega'] = pd.to_datetime(data['FechaEntrega'], format="%d-%m-%Y")


            #replace commas for numeric values
            data['Volumen'] = data['Volumen'].astype(str)
            data['Volumen'] = data.Volumen.str.strip() 
            data['Volumen'] = data['Volumen'].str.replace(',','.')
            data['Volumen'] = data['Volumen'].str.replace('*','')

            #filter out dates that are not current date
            data = data[(data['FechaEntrega'] == now_date)] 
        
            #datatype setup
            
            data['Centro'] = data['Centro'].astype(str)
            data['Volumen'] = data['Volumen'].astype(float)
            data['Pedido'] = data['Pedido'].astype(str)
            data['Posicion'] = data['Posicion'].astype(int)
            data['Equipo'] = data['Equipo'].astype(str)
            data['Placa'] = data['Placa'].astype(str)
            data['Entrega'] = data['Entrega'].astype(str)
            data['Cliente'] = data['Cliente'].astype(str)
            data['Obra'] = data['Obra'].astype(str)
            data['Denominacion'] = data['Denominacion'].astype(str)
            data['Estatus'] = data['Estatus'].astype(str)
                
            #strip string columns
            data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            data = data.replace(r'^\s*$', np.nan, regex=True)

            #merge factory data with data to get name of factory
            data = pd.merge(data,factories,on="Centro",how="left")
            data['Volumen'] = data['Volumen'].fillna(0)

            data.loc[data['Volumen'] > 1000, 'Volumen'] = data['Volumen'] / 1000


            #convert yards to m3 for factories that apply. Conversion rate 0.764555
            data.loc[data['NombrePlanta'].str.contains('PR-',na=False), 'Volumen'] = data['Volumen'] * 0.764555
            
            #now that we converted yards to m3 for the corresponding factories, drop the name of the factory column
            data = data.drop(['NombrePlanta'], axis=1) 
                
        return data
    except Exception as e:
        print(str(e))
     