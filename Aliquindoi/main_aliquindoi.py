import pandas as pd
import xlwings as xw
from openpyxl import Workbook
import traceback
from datetime import datetime
import os # Necesario para la ruta de los archivos

from programas import lectura_datos, plantillas_excel
from muestra import Muestra
from Aliquindoi.para_el_usuario.Configuracion.importar_configuracion import Config

config = Config()

try:
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

    datos_basicos = lectura_datos.pregunta_tipos_test()
    print("datos basicos: ",datos_basicos)

    # declaracion de variables antes de iniciar el programa.
    archivos_ir = archivos_zero_base_uv = archivos_muestra_uv = referencias_uv = referencias_ir = path_ftir = path_espectofotometro = ""

    # leer datos espectofotómetro/ftir si los hay
    if "FTIR" in datos_basicos["aparatos"]:
        path_ftir = lectura_datos.carpeta_con_archivos("Selecciona la carpeta ftir con las muestras y las referencias")
        print("path ftir seleccionada: ", path_ftir)
        if datos_basicos["aparatos"]["FTIR"] == "Con ventana":
            ventana_ftir = True
        else:
            ventana_ftir = False

        # elegir pestañas de referencias
        if "FTIR" in datos_basicos["aparatos"]:
            referencias_ir = {"r_ftir": "", "r_trans_ir": ""}
            r_ftir = lectura_datos.elegir_columnas_referencia("absorbedores_refl", "Selecciona referencia ftir")
            referencias_ir["r_ftir"] = r_ftir
            if ventana_ftir == True:
                r_trans_ir = lectura_datos.elegir_columnas_referencia("T_ventana_ir",
                                                                      "Selecciona transmitancia de la ventana")
                referencias_ir["r_trans_ir"] = r_trans_ir
            print("referencias ir seleccionadas: ", referencias_ir)


    if "Espectrofotómetro" in datos_basicos["aparatos"]:
        path_espectofotometro = lectura_datos.carpeta_con_archivos("Selecciona la carpeta de datos del espectrofotómetro")
        if datos_basicos["aparatos"]["Espectrofotómetro"] == "Con ventana":
            ventana_esp = True
        else:
            ventana_esp = False

        if "Espectrofotómetro" in datos_basicos["aparatos"]:
            referencias_uv = {"r_uv": "", "r_trans_uv": ""}
            if datos_basicos["medida"] == "Absortancia":
                r_uv = lectura_datos.elegir_columnas_referencia("absorbedores_abs", "Selecciona referencia base UV")
            elif datos_basicos["medida"] == "Reflectancia":
                r_uv = lectura_datos.elegir_columnas_referencia("reflectores", "Selecciona referencia base UV")
            elif (datos_basicos["medida"] == "Transmitancia CSP"):
                r_uv = lectura_datos.elegir_columnas_referencia("ref_trans_csp", "Selecciona referencia base UV")
            else:
                r_uv = ""
            referencias_uv["r_uv"] = r_uv

            if (datos_basicos["medida"] == "Absortancia") | (datos_basicos["medida"] == "Reflectancia"):
                if ventana_esp == True:
                    print("Ventana en UV seleccionada")
                    r_trans_uv = lectura_datos.elegir_columnas_referencia("T_ventana_uv",
                                                                          "Selecciona transmitancia de la ventana")
                    referencias_uv["r_trans_uv"] = r_trans_uv
            print("Referencias uv seleccionadas")
            print(referencias_uv)

    nombres_muestras = lectura_datos.nombres_muestras_auto(path_espectofotometro, path_ftir)
    print("nombres de muestras: ", nombres_muestras)
    instancias = {} #diccionario para guardar las instancias de las muestras

    for nombre in nombres_muestras:
        nombre_procesado = lectura_datos.procesar_cadena(nombre)
        print(f"Analizando muestra {nombre}")
        print(f"Nombre de la muestra procesado: {nombre_procesado}")

        if "FTIR" in datos_basicos["aparatos"]:
            # leer ftir si hay: path de reforo, zero, medidas (+refnegro, ventana, ventanaoro, ventananegro)
            archivos_ir = lectura_datos.ftir_medidas_auto(path_ftir, nombre_procesado, ventana_ftir)  # cambiar esta función

        if "Espectrofotómetro" in datos_basicos["aparatos"]:
            # leer carpeta de espectofotómetro si hay: path zero, base, muestras, (ventana y ventana base)
            archivos_zero_base_uv, archivos_muestra_uv = lectura_datos.espectro_medidas_zero_base_auto(nombre,
                                                                                                       path_espectofotometro,
                                                                                                       datos_basicos)
            print("archivos zero y base: ", archivos_zero_base_uv)
            print("archivos muestra: ", archivos_muestra_uv)

        #meter datos en la muestra
        print("Creando muestra")
        instancias[nombre] = Muestra(nombre, archivos_ir, archivos_zero_base_uv, archivos_muestra_uv, referencias_ir, referencias_uv, datos_basicos,
                                     excel_path_output)
        print("Muestra guardada: ", instancias[nombre])
        #Hacer el proceso de lectura de datos en la muestra
        if datos_basicos["medida"] == "Reflectancia":
            print("Reflectancia")
            plantillas_excel.copiar_datos_excel_config(instancias[nombre], wb_destino, config)
        elif datos_basicos["medida"] == "Absortancia":
            print("Absortancia")
            abs_ref_ir = pd.DataFrame()
            abs_ref_uv = pd.DataFrame()
            SWR_uv = SWA_uv = emitancia = ""
            dataframe_ir = pd.DataFrame() #si no hay ftir, se queda vacio
            if "FTIR" in datos_basicos["aparatos"] and archivos_ir:
                dataframe_ir = instancias[nombre].procesar_datos_tfir()
                abs_ref_ir = instancias[nombre].medidas_ir(dataframe_ir, ventana_ftir)
            if "Espectrofotómetro" in datos_basicos["aparatos"]:
                dataframe_uv = instancias[nombre].leer_datos_UV(ventana_esp)
                if dataframe_uv.shape[1]>2:
                    abs_ref_uv, SWR_uv, SWA_uv, SWR_std = instancias[nombre].medidas_UV(dataframe_uv, ventana_esp)
            else:
                abs_ref_uv= dataframe_uv = pd.DataFrame()
                SWR_uv = SWA_uv = SWR_std = ""
            data_absorbedor = instancias[nombre].combinar_uv_ir(abs_ref_ir, abs_ref_uv)
            if ("FTIR" in datos_basicos["aparatos"]) and ("Espectrofotómetro" in datos_basicos["aparatos"]) and archivos_ir:
                emitancia, df_abs = instancias[nombre].emitancia(data_absorbedor)
            else:
                emitancia = "no calculada" # no calcular la emitancia si falta algún rango
                df_abs = pd.DataFrame()
            plantillas_excel.copiar_datos_excel_absorbedores_config(instancias[nombre], data_absorbedor,wb_destino,
                                                             SWR_uv, SWA_uv, SWR_std, emitancia, instancias[nombre].temperatura,
                                                             dataframe_ir, dataframe_uv, df_abs, config)

        else:
            plantillas_excel.copiar_datos_excel(instancias[nombre], wb_destino, config)
            print("Transmitancia")

except Exception as e:
    # --- NUEVA LÓGICA DE REGISTRO DE ERRORES ---
    error_msg = traceback.format_exc()
    log_filename = excel_path_output.replace("xlsx", "txt")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Escribir el error en el archivo de log
    try:
        with open(log_filename, "a", encoding="utf-8") as f:
            f.write("=" * 80 + "\n")
            f.write(f"--- ERROR CRÍTICO OCURRIDO EN: {timestamp} ---\n")
            f.write("=" * 80 + "\n")
            f.write(error_msg)
            f.write("\n" + "=" * 80 + "\n\n")

        # Mensaje para el usuario en la consola
        print("\n" * 2)
        print("******************************************************************")
        print("!! ERROR CRÍTICO !!")
        print(f"Ocurrió un error inesperado durante la ejecución del programa.")
        print(f"Por favor, envíe el archivo de registro '{log_filename}' para su análisis.")
        print(f"Ruta del archivo: {os.path.abspath(log_filename)}")
        print("******************************************************************")
        print("\n")

    except IOError:
        # Esta excepción es la que el usuario está viendo el error de sintaxis.
        # Se ha reescrito para asegurar que esté indentada correctamente con respecto al 'try'
        print(
            f"\n[ERROR FATAL] Ocurrió un error inesperado, pero NO se pudo escribir en el archivo de log '{log_filename}'.")
        print(f"Detalle del error: {e}")

finally:
    # Intentar restaurar la configuración de Excel y cerrar el libro solo si se abrieron
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
            # Esto puede ocurrir si xlwings pierde la conexión al libro
            print(
                f"ADVERTENCIA: No se pudo guardar/cerrar el libro de destino. Por favor, cierre Excel manualmente. {clean_up_e}")
