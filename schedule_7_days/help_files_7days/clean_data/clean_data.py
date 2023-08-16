import pandas as pd
import numpy as np
import sys
import os
from help_files_7days.send_email.send_email import send_email

def clean_data(now,now_date,now_date1,wb_path):
    try:
        print("Cleaning data")
        data = pd.read_excel(wb_path,sheet_name="programacion")
        #rename columns to sql columns
        
        data = data.rename(columns={
                data.columns[0]: 'Entrega', 
                data.columns[1]: 'Remision', 
                data.columns[2]: 'Planta', 
                data.columns[3]: 'NombrePlanta', 
                data.columns[4]: 'Pedido',
                data.columns[5]: 'EstatusPedido', 
                data.columns[6]: 'VolumenPedido', 
                data.columns[7]: 'servicio',
                data.columns[8]: 'Posicion', 
                data.columns[9]: 'EstatusPosicion', 
                data.columns[10]: 'EstatusPosicion2',
                data.columns[11]: 'VolPartida', 
                data.columns[12]: 'Frecuencia', 
                data.columns[13]: 'Estatus',
                data.columns[14]: 'NombreCliente', 
                data.columns[15]: 'Cliente', 
                data.columns[16]: 'NombreObra',
                data.columns[17]: 'Obra', 
                data.columns[18]: 'NombreFrente', 
                data.columns[19]: 'Frente',
                data.columns[20]: 'ProductoComercial', 
                data.columns[21]: 'DescTecnica', 
                data.columns[22]: 'Solicitante',
                data.columns[23]: 'ComentarioInterno', 
                data.columns[24]: 'IDConductor', 
                data.columns[25]: 'NombreConductor',
                data.columns[26]: 'FechaReq', 
                data.columns[27]: 'HrReq', 
                data.columns[28]: 'FReqUlt',
                data.columns[29]: 'HrReqUlt', 
                data.columns[30]: 'FechaEntrega', 
                data.columns[31]: 'HoraEntregaPartida',
                data.columns[32]: 'CreadoEl', 
                data.columns[33]: 'HoraCreacion', 
                data.columns[34]: 'HRealInicioCarga',
                data.columns[35]: 'HRealFinCarga', 
                data.columns[36]: 'RHaciaObra', 
                data.columns[37]: 'REnObra',
                data.columns[38]: 'RHaciaPlanta', 
                data.columns[39]: 'RLlegada', 
                data.columns[40]: 'FechaCancelacion',
                data.columns[41]: 'HoraCancelacion', 
                data.columns[42]: 'ClaseDocumento', 
                data.columns[43]: 'Placa',
                data.columns[44]: 'ElementoAColar', 
                data.columns[45]: 'Contrato', 
                data.columns[46]: 'OrdenCompra',
                data.columns[47]: 'MtvCancelacion', 
                data.columns[48]: 'Latitud', 
                data.columns[49]: 'Longitud',
                data.columns[50]: 'Comercial', 
                data.columns[51]: 'HoraConsulta'})

        
        #covert datetime for date columns
        data['FechaReq'] = data['FechaReq'].astype(str)
        data['FechaReq'] = data['FechaReq'].str.replace('.','-')
        data['FechaReq'] = data['FechaReq'].replace(r'^\s*$', np.nan, regex=True)
        data['FechaReq'] = data['FechaReq'].astype(str)
        data['FechaReq'] = data.FechaReq.str.strip()
        data['FechaReq'] = pd.to_datetime(data['FechaReq'], format="%d-%m-%Y")

        data['FReqUlt'] = data['FReqUlt'].astype(str)
        data['FReqUlt'] = data['FReqUlt'].str.replace('.','-')
        data['FReqUlt'] = data['FReqUlt'].replace(r'^\s*$', np.nan, regex=True)
        data['FReqUlt'] = data['FReqUlt'].astype(str)
        data['FReqUlt'] = data.FReqUlt.str.strip()
        data['FReqUlt'] = pd.to_datetime(data['FReqUlt'], format="%d-%m-%Y")

        data['FechaEntrega'] = data['FechaEntrega'].astype(str)
        data['FechaEntrega'] = data['FechaEntrega'].str.replace('.','-')
        data['FechaEntrega'] = data['FechaEntrega'].replace(r'^\s*$', np.nan, regex=True)
        data['FechaEntrega'] = data['FechaEntrega'].astype(str)
        data['FechaEntrega'] = data.FechaEntrega.str.strip()
        data['FechaEntrega'] = pd.to_datetime(data['FechaEntrega'], format="%d-%m-%Y")

        data['CreadoEl'] = data['CreadoEl'].astype(str)
        data['CreadoEl'] = data['CreadoEl'].str.replace('.','-')
        data['CreadoEl'] = data['CreadoEl'].replace(r'^\s*$', np.nan, regex=True)
        data['CreadoEl'] = data['CreadoEl'].astype(str)
        data['CreadoEl'] = data.CreadoEl.str.strip()
        data['CreadoEl'] = pd.to_datetime(data['CreadoEl'], format="%d-%m-%Y")

        data['FechaCancelacion'] = data['FechaCancelacion'].astype(str)
        data['FechaCancelacion'] = data['FechaCancelacion'].str.replace('.','-')
        data['FechaCancelacion'] = data['FechaCancelacion'].replace(r'^\s*$', np.nan, regex=True)
        data['FechaCancelacion'] = data['FechaCancelacion'].astype(str)
        data['FechaCancelacion'] = data.FechaCancelacion.str.strip()
        data['FechaCancelacion'] = pd.to_datetime(data['FechaCancelacion'], format="%d-%m-%Y")

        #replace commas for numeric values
        data['VolumenPedido'] = data['VolumenPedido'].astype(str) 
        data['VolumenPedido'] = data.VolumenPedido.str.strip()
        data['VolumenPedido'] = data['VolumenPedido'].str.replace('.','')
        data['VolumenPedido'] = data['VolumenPedido'].str.replace(',','.')
        data['VolumenPedido'] = data['VolumenPedido'].str.replace('*','')

        #replace commas for numeric values
        data['VolPartida'] = data['VolPartida'].astype(str) 
        data['VolPartida'] = data.VolPartida.str.strip()
        data['VolPartida'] = data['VolPartida'].str.replace('.','')
        data['VolPartida'] = data['VolPartida'].str.replace(',','.')
        data['VolPartida'] = data['VolPartida'].str.replace('*','')
        
        #set lat and long tu null for now
        data['Latitud'] = None
        data['Longitud'] = None        

        #add timestamp column and reverse hora consulta and hora carga prog
        data['HoraCargaProg'] = data['HoraConsulta']
        data['HoraConsulta'] = now

        #filter out dates that are not current date
        data = data[(data['FechaEntrega'] >= now_date) & (data['FechaEntrega'] <= now_date1)]   

        #datatype setup
        
        data['Entrega'] = data['Entrega'].astype(str)
        data['Remision'] = data['Remision'].astype(str)
        data['Planta'] = data['Planta'].astype(str)
        data['NombrePlanta'] = data['NombrePlanta'].astype(str)
        data['Pedido'] = data['Pedido'].astype(str)
        data['EstatusPedido'] = data['EstatusPedido'].astype(str)
        data['VolumenPedido'] = data['VolumenPedido'].astype(float)
        data['servicio'] = data['servicio'].astype(str)
        data['Posicion'] = data['Posicion'].astype(int)
        data['EstatusPosicion'] = data['EstatusPosicion'].astype(str)
        data['EstatusPosicion2'] = data['EstatusPosicion2'].astype(str)
        data['VolPartida'] = data['VolPartida'].astype(float)
        data['Frecuencia'] = data['Frecuencia'].astype(int)
        data['Estatus'] = data['Estatus'].astype(str)
        data['NombreCliente'] = data['NombreCliente'].astype(str)
        data['Cliente'] = data['Cliente'].astype(int)
        data['NombreObra'] = data['NombreObra'].astype(str)
        data['Obra'] = data['Obra'].astype(str)
        data['NombreFrente'] = data['NombreFrente'].astype(str)
        data['Frente'] = data['Frente'].astype(str)
        data['ProductoComercial'] = data['ProductoComercial'].astype(int)
        data['DescTecnica'] = data['DescTecnica'].astype(str)
        data['Solicitante'] = data['Solicitante'].astype(str)
        data['ComentarioInterno'] = data['ComentarioInterno'].astype(str)
        data['IDConductor'] = data['IDConductor'].astype(str)
        data['NombreConductor'] = data['NombreConductor'].astype(str)
        data['ClaseDocumento'] = data['ClaseDocumento'].astype(str)
        data['Placa'] = data['Placa'].astype(str)
        data['ElementoAColar'] = data['ElementoAColar'].astype(str)
        data['Contrato'] = data['Contrato'].astype(str)
        data['OrdenCompra'] = data['OrdenCompra'].astype(str)
        data['MtvCancelacion'] = data['MtvCancelacion'].astype(str)
        data['Comercial'] = data['Comercial'].astype(str)
        data['HrReq'] = data['HrReq'].astype(str)
        data['HrReqUlt'] = data['HrReqUlt'].astype(str)
        data['HoraEntregaPartida'] = data['HoraEntregaPartida'].astype(str)
        data['HoraCreacion'] = data['HoraCreacion'].astype(str)
        data['HRealInicioCarga'] = data['HRealInicioCarga'].astype(str)
        data['HRealFinCarga'] = data['HRealFinCarga'].astype(str)
        data['RHaciaObra'] = data['RHaciaObra'].astype(str)
        data['REnObra'] = data['REnObra'].astype(str)
        data['RHaciaPlanta'] = data['RHaciaPlanta'].astype(str)
        data['RLlegada'] = data['RLlegada'].astype(str)
        data['HoraCancelacion'] = data['HoraCancelacion'].astype(str)
        data['HoraCargaProg'] = data['HoraCargaProg'].astype(str)
        
        #strip string columns
        data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        #replace string blanks for null values
        data = data.replace(r'^\s*$', np.nan, regex=True)

        data['VolPartida'] = data['VolPartida'].fillna(0)
        data['VolumenPedido'] = data['VolumenPedido'].fillna(0)

        data.loc[data['VolPartida'] >= 1000, 'VolPartida'] = data['VolPartida'] / 1000
        data.loc[data['VolumenPedido'] >= 1000, 'VolumenPedido'] = data['VolumenPedido'] / 1000
        
        #convert yards to m3 for factories that apply. Conversion rate 0.764555
        data.loc[data['NombrePlanta'].str.contains('PR-',na=False), 'VolPartida'] = data['VolPartida'] * 0.764555
        data.loc[data['NombrePlanta'].str.contains('PR-',na=False), 'VolumenPedido'] = data['VolumenPedido'] * 0.764555
        

        return data
    except Exception as e:
        print(str(e))
        send_email(str(e))
        sys.exit()

def clean_data2(now,now_date,now_date1,path,filename):
    try:
        print("Cleaning data2")
        full_path = os.path.join(path,filename)
        data = pd.read_csv(full_path,sep='|', encoding='latin-1', on_bad_lines='skip', skiprows=3)
        
        #get total rows in dataframe to then filter out the first and last row that are non valid rows
        data_len = len(data.index)-2

        data = data.loc[1:data_len]

        #clean dataset
        data = data.reset_index(drop=True)

        #drop unnamed columns
        data = data.drop(columns=['Unnamed: 0'])
        data = data.drop(columns=['Unnamed: '+str(len(data.columns.values))]) 
        
        data = data.rename(columns={
                data.columns[0]: 'FechaEntrega', 
                data.columns[1]: 'Pedido', 
                data.columns[2]: 'enlistaespera',
                data.columns[3]: 'estatus'})
        
        
        #covert datetime for date columns
        
        data['FechaEntrega'] = data['FechaEntrega'].astype(str)
        data['FechaEntrega'] = data['FechaEntrega'].str.replace('.','-')
        data['FechaEntrega'] = data['FechaEntrega'].replace(r'^\s*$', np.nan, regex=True)
        data['FechaEntrega'] = data['FechaEntrega'].astype(str)
        data['FechaEntrega'] = data.FechaEntrega.str.strip()
        data['FechaEntrega'] = pd.to_datetime(data['FechaEntrega'], format="%d-%m-%Y")

        data = data[data['estatus'].str.contains('Por confirmar')]
        
        data = data[data['enlistaespera'].str.contains('X')]

        data = data.drop(columns=['estatus'])
        
        

        #filter out dates that are not current date
        data = data[(data['FechaEntrega'] >= now_date) & (data['FechaEntrega'] <= now_date1)]   

        #datatype setup
        
        
        data['Pedido'] = data['Pedido'].astype(str)
        data['enlistaespera'] = data['enlistaespera'].astype(str)
        
        
        #strip string columns
        data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        #replace string blanks for null values
        data = data.replace(r'^\s*$', np.nan, regex=True)
        
     
        
        
        return data
    except Exception as e:
        print(str(e))
        send_email(str(e))
        sys.exit()
