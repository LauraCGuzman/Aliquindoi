from programas import lectura_datos
from muestra import Muestra
import numpy as np

respuesta = True

while respuesta == True:
    #meter nombre de todas las muestras
    # cuadro de dialogo para meter todas las muestras y automatizar el proceso
    datos_basicos = lectura_datos.pregunta_tipos_test()
    print(datos_basicos)
    # introducir nombres de las muestras
    nombres_muestras = lectura_datos.nombres_muestras_auto()
    # leer datos espectofotómetro/ftir si los hay
    if "FTIR" in datos_basicos["aparatos"]:
        path_ftir_muestras = lectura_datos.archivo_tfir("Selecciona archivo FTIR de las muestras \n o muestras y referencias")
        path_ftir_referencias = lectura_datos.archivo_tfir("Selecciona archivo FTIR de las referencias")

        if datos_basicos["aparatos"]["FTIR"] == "Con ventana":
            ventana_ftir = True
        else:
            ventana_fitr = False

    if "Espectrofotómetro" in datos_basicos["aparatos"]:
        path_espectofotometro = lectura_datos.carpeta_espectofotometro()

    #preguntar por referencias uv/ir:
    instancias = {} #diccionario para guardar las instancias de las muestras
    for nombre in nombres_muestras:
        
        print(f"Analizando muestra {nombre}")
        #elegir pestañas de las muestras y referencias
        if "FTIR" in datos_basicos["aparatos"]:
            archivos_ir = {} #declarar diccionario para guardar los archivos de ftir
            if path_ftir_muestras and path_ftir_referencias:
                #leer ftir si hay: path de reforo, zero, medidas (+refnegro, ventana, ventanaoro, ventananegro)
                archivos_ir = lectura_datos.ftir_medidas_auto(archivos_ir, path_ftir_muestras, ventana_ftir, "muestras") #cambiar esta función
                archivos_ir = lectura_datos.ftir_medidas_auto(archivos_ir, path_ftir_referencias, ventana_ftir, "referencias")
            else:
                archivos_ir = lectura_datos.ftir_medidas_auto(archivos_ir, path_ftir_muestras, ventana_ftir, "ambos")
            #referencias ftir <- Falta
            r_oro_ir = lectura_datos.elegir_columnas_referencia("absorbedores_refl", "Selecciona referencia oro")
            if ventana_ftir == True:
                r_fsd_ir = lectura_datos.elegir_columnas_referencia("absorbedores_abs", "Selecciona referencia negro")
                r_trans_ir = lectura_datos.elegir_columnas_referencia("T_ventana", "Selecciona transmitancia de la ventana")

        
        if "Espectrofotómetro" in datos_basicos["aparatos"]:
            #leer carpeta de espectofotómetro si hay: path zero, base, muestras, (ventana y ventana base)
            archivos_zero_base_uv, archivos_muestra_uv = lectura_datos.espectro_medidas_zero_base_auto(nombre, path_espectofotometro)
            #referencias uv <- Falta
            r_fsd_ir = lectura_datos.elegir_columnas_referencia("absorbedores_abs", "Selecciona referencia base")
            #falta añadir a las referencias la transmitancia de la ventana en UV

        
        #meter datos en la muestra
        #instancias{nombre} = Muestra(nombre, archivos_muestra_ir, archivos_zero_base_uv, archivos_muestra_uv, referncias_uv, referencias_ir, datos_basicos)
        
        #Hacer el proceso de lectura de datos en la muestra
        if datos_basicos["medida"] == "Reflectancia":
            print("Reflectancia")
        elif datos_basicos["medida"] == "Absortancia":
            print("Absortancia")
        elif datos_basicos["medida"] == "Transmitancia":
            print("Transmitancia")
    
    respuesta = lectura_datos.detener_analisis() #pregunta al usuario si quiere seguir analizando muestras