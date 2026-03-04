import sys
import os

# Add project root to sys.path to allow imports from Aliquindoi and user_templates
# Assumes this file is in <project_root>/Aliquindoi/
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import pandas as pd
import xlwings as xw
from openpyxl import Workbook
import traceback
from datetime import datetime
import os # Necesario para la ruta de los archivos

from Aliquindoi.programas import lectura_datos, plantillas_excel
from Aliquindoi.muestra import Muestra
from user_templates.Configuracion.importar_configuracion import Config

def setup_application():
    """Configura el archivo Excel de salida y la aplicación."""
    # cuadro de dialogo para meter todas las muestras y automatizar el proceso
    excel_path_output = lectura_datos.preguntar_output_excel()
    print(excel_path_output)
    workbook = Workbook()
    workbook.save(excel_path_output)

    wb_destino = xw.Book(excel_path_output)  # Abrir el workbook una sola vez
    # desactivar funciones para ir más rápido
    app = wb_destino.app  # Obtener la aplicación de Excel
    app.screen_updating = False  # Evitar parpadeos en pantalla
    app.calculation = 'manual'  # Desactivar el cálculo automático
    
    return excel_path_output, wb_destino, app

def configure_references_ftir(datos_basicos):
    """Configura las referencias y rutas para FTIR si está seleccionado."""
    path_ftir = ""
    ventana_ftir = False
    referencias_ir = ""

    if "FTIR" in datos_basicos["aparatos"]:
        path_ftir = lectura_datos.carpeta_con_archivos("Selecciona la carpeta ftir con las muestras y las referencias")
        print("path ftir seleccionada: ", path_ftir)
        
        if datos_basicos["aparatos"]["FTIR"] == "Con ventana":
            ventana_ftir = True
        
        # elegir pestañas de referencias
        referencias_ir = {"r_ftir": "", "r_trans_ir": ""}
        r_ftir = lectura_datos.elegir_columnas_referencia("absorbedores_refl", "Selecciona referencia ftir")
        referencias_ir["r_ftir"] = r_ftir
        
        if ventana_ftir:
            r_trans_ir = lectura_datos.elegir_columnas_referencia("T_ventana_ir",
                                                                  "Selecciona transmitancia de la ventana")
            referencias_ir["r_trans_ir"] = r_trans_ir
        print("referencias ir seleccionadas: ", referencias_ir)
            
    return path_ftir, ventana_ftir, referencias_ir

def configure_references_uv(datos_basicos):
    """Configura las referencias y rutas para Espectrofotómetro si está seleccionado."""
    path_espectofotometro = ""
    ventana_esp = False
    referencias_uv = ""

    if "Espectrofotómetro" in datos_basicos["aparatos"]:
        path_espectofotometro = lectura_datos.carpeta_con_archivos("Selecciona la carpeta de datos del espectrofotómetro")
        
        if datos_basicos["aparatos"]["Espectrofotómetro"] == "Con ventana":
            ventana_esp = True
        
        # Solo solicitar referencias para Absortancia y Reflectancia
        # Transmitancia CSP/PV no requieren referencias ya que solo copian datos directamente
        if datos_basicos["medida"] in ["Absortancia", "Reflectancia"]:
            referencias_uv = {"r_uv": "", "r_trans_uv": ""}
            
            if datos_basicos["medida"] == "Absortancia":
                r_uv = lectura_datos.elegir_columnas_referencia("absorbedores_abs", "Selecciona referencia base UV")
            elif datos_basicos["medida"] == "Reflectancia":
                r_uv = lectura_datos.elegir_columnas_referencia("reflectores", "Selecciona referencia base UV")
            
            referencias_uv["r_uv"] = r_uv

            if ventana_esp:
                print("Ventana en UV seleccionada")
                r_trans_uv = lectura_datos.elegir_columnas_referencia("T_ventana_uv",
                                                                      "Selecciona transmitancia de la ventana")
                referencias_uv["r_trans_uv"] = r_trans_uv
            print("Referencias uv seleccionadas")
            print(referencias_uv)
        else:
            # Para Transmitancia CSP/PV, no se solicitan referencias
            referencias_uv = {"r_uv": "", "r_trans_uv": ""}
            print("Transmitancia CSP/PV seleccionada: no se requieren referencias")
        
    return path_espectofotometro, ventana_esp, referencias_uv

# Main Execution Flow
config = Config()
wb_destino = None
app = None

try:
    # 1. Setup Application
    excel_path_output, wb_destino, app = setup_application()
    
    # 2. Get User Input
    datos_basicos = lectura_datos.pregunta_tipos_test()
    print("datos basicos: ", datos_basicos)

    # 3. Configure References
    path_ftir, ventana_ftir, referencias_ir = configure_references_ftir(datos_basicos)
    path_espectofotometro, ventana_esp, referencias_uv = configure_references_uv(datos_basicos)

    # 4. Discover Samples
    nombres_muestras = lectura_datos.nombres_muestras_auto(path_espectofotometro, path_ftir)
    print("nombres de muestras: ", nombres_muestras)
    instancias = {} 

    # 5. Process Loops (Explicit Orchestration)
    for nombre in nombres_muestras:
        # Initialize loop-specific variables
        archivos_ir = ""
        archivos_zero_base_uv = ""
        archivos_muestra_uv = ""
        
        nombre_procesado = lectura_datos.procesar_cadena(nombre)
        print(f"Analizando muestra {nombre}")
        print(f"Nombre de la muestra procesado: {nombre_procesado}")

        # Resolve files based on apparatus
        if "FTIR" in datos_basicos["aparatos"]:
            archivos_ir = lectura_datos.ftir_medidas_auto(path_ftir, nombre_procesado, ventana_ftir)

        if "Espectrofotómetro" in datos_basicos["aparatos"]:
            archivos_zero_base_uv, archivos_muestra_uv = lectura_datos.espectro_medidas_zero_base_auto(nombre,
                                                                                                       path_espectofotometro,
                                                                                                       datos_basicos)
            print("archivos zero y base: ", archivos_zero_base_uv)
            print("archivos muestra: ", archivos_muestra_uv)

        # Create Domain Object
        print("Creando muestra")
        instancias[nombre] = Muestra(nombre, archivos_ir, archivos_zero_base_uv, archivos_muestra_uv, referencias_ir, referencias_uv, datos_basicos,
                                     excel_path_output)
        print("Muestra guardada: ", instancias[nombre])
        
        # Calculate and Report based on Measurement Type
        if datos_basicos["medida"] == "Reflectancia":
            print("Reflectancia")
            plantillas_excel.copiar_datos_excel(instancias[nombre], wb_destino, config)
            
        elif datos_basicos["medida"] == "Absortancia":
            print("Absortancia")
            abs_ref_ir = pd.DataFrame()
            abs_ref_uv = pd.DataFrame()
            SWR_uv = SWA_uv = SWR_std = emitancia = ""
            dataframe_ir = pd.DataFrame() 
            dataframe_uv = pd.DataFrame() # Add explicit init for safety
             
            if "FTIR" in datos_basicos["aparatos"] and archivos_ir:
                dataframe_ir = instancias[nombre].procesar_datos_tfir()
                abs_ref_ir = instancias[nombre].medidas_ir(dataframe_ir, ventana_ftir)
                
            if "Espectrofotómetro" in datos_basicos["aparatos"]:
                dataframe_uv = instancias[nombre].leer_datos_UV(ventana_esp)
                if dataframe_uv.shape[1]>2:
                    abs_ref_uv, SWR_uv, SWA_uv, SWR_std = instancias[nombre].medidas_UV(dataframe_uv, ventana_esp)
            else:
                 # Logic from original else block
                abs_ref_uv = dataframe_uv = pd.DataFrame()
                SWR_uv = SWA_uv = SWR_std = ""
                
            data_absorbedor = instancias[nombre].combinar_uv_ir(abs_ref_ir, abs_ref_uv)
            
            if ("FTIR" in datos_basicos["aparatos"]) and ("Espectrofotómetro" in datos_basicos["aparatos"]) and archivos_ir:
                emitancia, df_abs = instancias[nombre].emitancia(data_absorbedor)
            else:
                emitancia = "no calculada"
                df_abs = pd.DataFrame()
                
            plantillas_excel.copiar_datos_excel_absorbedores(instancias[nombre], data_absorbedor, wb_destino,
                                                             SWR_uv, SWA_uv, SWR_std, emitancia, instancias[nombre].temperatura,
                                                             dataframe_ir, dataframe_uv, df_abs, config)

        else: # Transmitancia CSP or PV
            plantillas_excel.copiar_datos_excel(instancias[nombre], wb_destino, config)
            print("Transmitancia")

except Exception as e:
    # Error Handling Legacy Logic
    error_msg = traceback.format_exc()
    if 'excel_path_output' in locals() and excel_path_output:
        log_filename = excel_path_output.replace("xlsx", "txt")
    else:
        log_filename = "error_log.txt"
        
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        with open(log_filename, "a", encoding="utf-8") as f:
            f.write("=" * 80 + "\n")
            f.write(f"--- ERROR CRÍTICO OCURRIDO EN: {timestamp} ---\n")
            f.write("=" * 80 + "\n")
            f.write(error_msg)
            f.write("\n" + "=" * 80 + "\n\n")

        print("\n" * 2)
        print("******************************************************************")
        print("!! ERROR CRÍTICO !!")
        print(f"Ocurrió un error inesperado durante la ejecución del programa.")
        print(f"Por favor, envíe el archivo de registro '{log_filename}' para su análisis.")
        print(f"Ruta del archivo: {os.path.abspath(log_filename)}")
        print("******************************************************************")
        print("\n")

    except IOError:
        print(f"\n[ERROR FATAL] Ocurrió un error inesperado, pero NO se pudo escribir en el archivo de log '{log_filename}'.")
        print(f"Detalle del error: {e}")

finally:
    # Cleanup
    if app is not None:
        try:
            app.screen_updating = True
            app.calculation = 'automatic'
        except Exception as clean_up_e:
            print(f"ADVERTENCIA: No se pudo restaurar la configuración de la aplicación Excel. {clean_up_e}")

    if wb_destino is not None:
        try:
            wb_destino.save()
            wb_destino.close()
        except Exception as clean_up_e:
            print(f"ADVERTENCIA: No se pudo guardar/cerrar el libro de destino. Por favor, cierre Excel manualmente. {clean_up_e}")
