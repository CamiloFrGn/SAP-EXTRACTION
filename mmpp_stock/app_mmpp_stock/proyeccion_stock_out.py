import pandas as pd
import sys
#sys.path.insert(0, r'C:\Users\jsdelgadoc\Documents\otros\mmpp_stock')
sys.path.insert(0, r'C:\Users\E-JFRANCOGON\Downloads\sap_extraction\mmpp_stock')
from help_files_mmpp_stock.sql.connect_sql_server import *

def main_proyeccion():
    try:    
        print("RUNNING PROGRAM FOR MMPP PROYECTION STOCK OUT")

        procedure_name = 'scac_ap_proyeccion_stockout_v3'
        merged_df = query_sql_df( 
                    "{CALL " + procedure_name+ " ()}", 
                    ( 
                    ) 
                )
        print(merged_df.head())
        # Merge the tables on the 'Material' column
        ##merged_df = pd.merge(daily_demand_df, current_stock_df, on='Material')

        # Convert the 'Date' column to datetime type
        merged_df['fecha'] = pd.to_datetime(merged_df['fecha'])

        # Get unique materials and dates
        materials = merged_df['Material'].unique()
        dates = merged_df['fecha'].unique()

        # Perform cross join between materials and dates
        ##cross_df = pd.DataFrame(list(itertools.product(materials, dates)), columns=['Material', 'Date'])

        # Merge cross_df with merged_df to get all combinations of materials and dates
        ##merged_df = pd.merge(cross_df, merged_df, on=['Material', 'Date'], how='left')

        # Sort the dataframe by 'Material' and 'Date'
        ##merged_df.sort_values(['Material', 'Date'], inplace=True)
        print("####################")

        merged_df['stock_actual'] = merged_df['stock_actual'].astype(float)
        merged_df['demanda'] = merged_df['demanda'].astype(float)

        # Create a new column 'Calculated_Stock' initialized with the 'Stock' values
        merged_df['calculated_Stock'] = merged_df['stock_actual']

        # Iterate over the rows and update 'Calculated_Stock' based on the conditions
        for i, row in merged_df.iterrows():
            if row['minfecha'] == "minfecha":
                merged_df.at[i, 'calculated_Stock'] -= row['demanda']
            else:
                prev_row = merged_df.loc[i-1]
                merged_df.at[i, 'calculated_Stock'] = prev_row['calculated_Stock'] - row['demanda']

        merged_df = merged_df.drop("minfecha",axis=1)

        databasename = 'scac_at_proyeccion_stockout'
        #delete data before insert
        delete_sql = "delete from "+databasename
        query_sql_crud(delete_sql,())
        #insert data
      
        send_df_to_sql(merged_df,databasename)

    except Exception as e:
        print(str(e))
        sys.exit()

