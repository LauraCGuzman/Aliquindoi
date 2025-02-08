import pandas as pd
import sqlite3


# Rellena estos datos (Fill in these details)
database_sql = "C:/Users/camp_la/Documents/Laura/OPAC/proyecto_automatizar_procesos/primera_base_de_datos.db"
excel_to_copy = "C:/Users/camp_la/Documents/Laura/OPAC/proyects/OPAC_TableReflectanceALL.xlsm"
excel_sheet_name = "ReflectorsALL"
sql_table_name = "OPAC_TableReflectanceALL"

# Read and transform excel table
def prepare_excel_fow_sql_table_database(excel_to_copy):
    wb_test = pd.read_excel(excel_to_copy, sheet_name=excel_sheet_name)
    wb_test.columns = [col.replace(' ', '_').lower() for col in wb_test.columns] #replace spaces by _ and lower letters
    wb_test = wb_test.rename(columns={'unnamed:_24': 'ρs,h_std', 'unnamed:_27': 'astm_drop_std', 'unnamed:_30': 'hemispherical_iso_ρs,h_std',
                                     'unnamed:_33': 'iso_drop_std', 'unnamed:_36': 'hemispherical_660_ρλ,h_std',
                                     'unnamed:_39': '660_drop_std', 'unnamed:_42': 'ρλ,φ(660nm,15°,12.5 mrad)_std',
                                     'unnamed:_45': 'specular_drop_std', 'blistering_level_paint_layer': 'blistering_level',
                                     'unnamed:_59': 'paint_layer'})
    wb_test = wb_test.drop(wb_test.columns[wb_test.columns.str.contains('unnamed', case=False)], axis=1)
    return wb_test
# Transform table into database

def transform_excel_into_sql_table_database(wb_test, database_sql):
    try:
        # Connect to SQLite database (check for write access)
        cxn = sqlite3.connect(database_sql)
        cursor = cxn.cursor()
    
        # Check if database is read-only (if so, prompt to change permissions)
        cursor.execute("PRAGMA quick_check")
        results = cursor.fetchall()

        expected_status = "read-write"
        if isinstance(results[0][0], str) and expected_status not in results[0][0]:
            print("WARNING: The database seems to be read-only. Please ensure write access to continue.")
            user_choice = input("Do you want to try connecting to a writable database (y/n)? ")
            if user_choice.lower() != 'y':
                raise sqlite3.OperationalError("Database write access required. Exiting.")

        # Write to table, handling existing table and index
        wb_test.to_sql(name=sql_table_name, con=cxn, if_exists='replace', index=True)

        # Commit changes
        cxn.commit()

        print("Data successfully written to SQLite database.")

    except sqlite3.OperationalError as e:
        print(f"Error connecting to database: {e}")
    except Exception as e:  # Catch other potential errors
        print(f"An error occurred: {e}")

    finally:
        # Always close the connection, even if errors occur
        if cxn:
            cxn.close()
            print("Connection closed.")