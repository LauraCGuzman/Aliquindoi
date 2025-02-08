from programas import lectura_datos
from muestra import Muestra
respuesta = True
import numpy as np

while respuesta == True:
    #meter nombre de todas las muestras
    # cuadro de dialogo para meter todas las muestras y automatizar el proceso
    datos_basicos = lectura_datos.pregunta_tipos_test()
    print(datos_basicos)
    # introducir nombres de las muestras
    nombres_muestras = lectura_datos.nombres_muestras_auto()
    # leer datos espectofotómetro/ftir si los hay
    if "FTIR" in datos_basicos["aparatos"]:
        print("leer datos ftir")
        path_ftir_muestras = lectura_datos.archivo_tfir("Selecciona archivo FTIR de las muestras \n o muestras y referencias")
        path_ftir_referencias = lectura_datos.archivo_tfir("Selecciona archivo FTIR de las referencias")

        if path_ftir_muestras and path_ftir_referencias:
            dos_tfir = True
        else:
            dos_tfir = False
    if "Espectrofotómetro" in datos_basicos["aparatos"]:
        print("leer datos espectofotómetro")
        path_espectofotometro = lectura_datos.carpeta_espectofotometro()

    #preguntar por referencias uv/ir:

    for nombre in nombres_muestras:
        print(f"Analizando muestra {nombre}")
        #leer ftir si hay: path de reforo, zero, medidas (+refnegro, ventana, ventanaoro, ventananegro)
        #referencias ftir
        #leer carpeta de especofotómetro si hay: path zero, base, muestras, (ventana y ventana base)
        #referencias uv
        #meter datos en la muestra
        if datos_basicos["medida"] == "Reflectancia":
            print("Reflectancia")
        elif datos_basicos["medida"] == "Absortancia":
            print("Absortancia")
        elif datos_basicos["medida"] == "Transmitancia":
            print("Transmitancia")
    #Elegir modo automatico/manual
    """if datos_basicos["mode"] == "Manual":
        print("Modo manual seleccionado")
        nombre_muestra = lectura_datos.nombre_muestra()
        if datos_basicos["item"] == "Absorbedor":
            info_tfir = lectura_datos.archivo_tfir_base_zero_samples()
        else:
            info_tfir = np.nan
        file_path_zero_base, file_paths_muestras = lectura_datos.archivos_espectofotometro()
        if datos_basicos["item"] == "Absorbedor":
            sel_spec_ref = lectura_datos.seleccionar_spectro_referencia_IR_UV()
        else:
            sel_spec_ref = lectura_datos.seleccionar_spectro_referencia_UV()

        # escribir datos de partida en la muestra
        muestra = Muestra(nombre_muestra, info_tfir, file_path_zero_base, file_paths_muestras,
                          sel_spec_ref["ir_var"], sel_spec_ref["uv_var"], datos_basicos)
        print("Muestra: ", muestra.nombre)
        print("Archivo excel tfir: ",
              muestra.archivo_tfir)  # Nota: oro y zeroline no tienen que estar en el mismo archivo
        print("Lista archivos espectofotómetro: ", muestra.lista_espect_muestras)
        print("Directorio baseline: ", muestra.path_base)
        print("Directorio zeroline: ", muestra.path_zero)
        print("Columna referencia UV: ", muestra.col_uv_ref)
        print("Columna referencia IR: ", muestra.col_ir_ref)

    elif datos_basicos["mode"] == "Automático":
        print("Modo automático seleccionado")
        nombres_muestras = lectura_datos.nombres_muestras_auto()
        if datos_basicos["item"] == "Absorbedor":
            file_ftir = lectura_datos.archivo_tfir()
        else:
            file_ftir = np.nan
        file_path_espectofotometro = lectura_datos.carpeta_espectofotometro()
        if datos_basicos["item"] == "Absorbedor":
            sel_spec_ref = lectura_datos.seleccionar_spectro_referencia_IR_UV()
        else:
            sel_spec_ref = lectura_datos.seleccionar_spectro_referencia_UV()
        for nombre_muestra in nombres_muestras:
            print(f"Analizando muestra {nombre_muestra}")
            info_tfir = lectura_datos.tfir_medidas_zero_base_auto(nombre_muestra, file_ftir)

            file_path_zero_base, file_paths_muestras = lectura_datos.espectro_medidas_zero_base_auto(nombre_muestra, file_path_espectofotometro)

            # escribir datos de partida en la muestra
            muestra = Muestra(nombre_muestra, info_tfir, file_path_zero_base, file_paths_muestras,
                              sel_spec_ref["ir_var"], sel_spec_ref["uv_var"], datos_basicos)
            print("Muestra: ", muestra.nombre)
            print("Archivo excel tfir: ",
                  muestra.archivo_tfir)  # Nota: oro y zeroline no tienen que estar en el mismo archivo
            print("Lista archivos espectofotómetro: ", muestra.lista_espect_muestras)
            print("Directorio baseline: ", muestra.path_base)
            print("Directorio zeroline: ", muestra.path_zero)
            print("Columna referencia UV: ", muestra.col_uv_ref)
            print("Columna referencia IR: ", muestra.col_ir_ref)
    #Fin de toma de datos
    #Procesamiento de datos
    data_ir = muestra.leer_datos_muestra_IR() #devuelve NaN si es un reflector.
    print("Datos IR")
    print(data_ir)
    data_uv, solar_w_ref, solar_w_abs = muestra.leer_datos_muestra_UV()
    print("Datos UV")
    print(data_uv)
    print("solar w ref: ", solar_w_ref, "solar w abs: ", solar_w_abs)
    data_all = muestra.combinar_uv_ir(data_ir,data_uv)
    print("Datos UV e IR")
    print(data_all)"""

    respuesta = lectura_datos.detener_analisis() #pregunta al usuario si quiere seguir analizando muestras