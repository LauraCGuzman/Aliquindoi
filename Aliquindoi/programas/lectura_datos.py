from tkinter import simpledialog
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import tkinter as Tk
from tkinter import messagebox
from tkinter import ttk
from tkcalendar import Calendar
import re
from pathlib import Path


def leer_archivo_referencias():
    directorio_script = os.path.dirname(os.path.realpath(__file__))
    archivo_referencias = os.path.normpath(
        os.path.join(directorio_script, "../user_templates/references.xlsx"))

    return archivo_referencias

def preguntar_output_excel():
    """Abre un cuadro de diálogo para que el usuario seleccione dónde guardar el archivo Excel."""
    root = Tk.Tk()
    root.withdraw()  # Oculta la ventana principal

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        title="Guardar archivo Excel como..."
    )
    root.update()  # Asegura que los eventos de Tkinter se procesen correctamente
    root.destroy()
    return file_path

def pregunta_tipos_test():
    # Crear un diccionario para almacenar las variables seleccionadas
    resultados = {"medida": None, "aparatos": {}, "test": None, "fabricante": None, "proyecto":None,
                  "hours": None, "months": None, "temperatura": None, "fecha_medida": {}}

    def update_ventana_state(aparato):
        """ Habilita o deshabilita las opciones 'con ventana' y 'sin ventana' según el estado del aparato """
        if variables_aparatos[aparato].get():
            radio_ventana_si[aparato].config(state=Tk.NORMAL)
            radio_ventana_no[aparato].config(state=Tk.NORMAL)
        else:
            radio_ventana_si[aparato].config(state=Tk.DISABLED)
            radio_ventana_no[aparato].config(state=Tk.DISABLED)
            variable_ventana[aparato].set("")  # Reiniciar selección de ventana si se desmarca el aparato

    def abrir_calendario():
        """Abre el calendario en una ventana emergente."""
        top = Tk.Toplevel(root)
        cal = Calendar(top, selectmode="day", date_pattern="dd/mm/yyyy")
        cal.pack(pady=10)

        def seleccionar_fecha():
            fecha_seleccionada = cal.get_date()
            fecha_dt = datetime.strptime(fecha_seleccionada, "%d/%m/%Y")
            fecha_formato1 = fecha_dt.strftime("%d/%m/%Y")
            fecha_formato2 = fecha_dt.strftime("%Y%m%d")
            resultados["fecha_medida"] = {"dd/mm/yyyy": fecha_formato1, "yyyyMMdd": fecha_formato2}
            label_mostrar_fecha.config(text=f"Fecha seleccionada: {fecha_formato1}")
            top.destroy()

        Tk.Button(top, text="Seleccionar", command=seleccionar_fecha).pack(pady=5)

    def submit():
        selected_medida = variable_medida.get()
        selected_test = variable_test.get()
        selected_fabricante = variable_fabricante.get()
        selected_proyecto = variable_proyecto.get()
        hours = entry_hours.get()
        months = entry_months.get()
        temperatura = entry_t.get()

        if not selected_medida or not selected_test or not hours.isdigit():
            messagebox.showwarning("Advertencia", "Por favor, completa todas las opciones correctamente.")
            return

        resultados["medida"] = selected_medida
        resultados["test"] = selected_test
        resultados["fabricante"] = selected_fabricante
        resultados["proyecto"] = selected_proyecto

        # Validar y almacenar horas si se ingresan
        if hours:
            if hours.isdigit():
                resultados["hours"] = int(hours)
            else:
                messagebox.showwarning("Advertencia", "Horas debe ser un número entero.")
                return
        else:
            resultados["hours"] = None  # Almacenar None si el campo está vacío

        # Validar y almacenar meses si se ingresan
        if months:
            if months.isdigit():
                resultados["months"] = int(months)
            else:
                messagebox.showwarning("Advertencia", "Meses debe ser un número entero.")
                return
        else:
            resultados["months"] = None  # Almacenar None si el campo está vacío

        # Validar y almacenar temperatura si se ingresa
        if temperatura:
            try:
                resultados["temperatura"] = float(temperatura)
            except ValueError:
                messagebox.showwarning("Advertencia", "Temperatura debe ser un número.")
                return
        else:
            resultados["temperatura"] = None  # Almacenar None si el campo está vacío

        # Obtener selección de aparatos y si tienen ventana o no
        for aparato in ["FTIR", "Espectrofotómetro"]:
            if variables_aparatos[aparato].get():
                resultados["aparatos"][aparato] = variable_ventana[aparato].get()

        if not resultados["aparatos"]:
            messagebox.showwarning("Advertencia", "Debes seleccionar al menos un aparato.")
            return

        # Verificar si se seleccionó una fecha
        if not resultados["fecha_medida"]:
            messagebox.showwarning("Advertencia", "Por favor, selecciona una fecha.")
            return

        messagebox.showinfo("Selección", f"Has seleccionado:\nMedida: {selected_medida}\n"
                                         f"Aparatos: {resultados['aparatos']}\n"
                                         f"Tipo de test: {selected_test}\n"
                                         f"Fabricante: {selected_fabricante}\n"
                                         f"Proyecto: {selected_proyecto}\n"
                                         f"Horas: {hours}\n"
                                         f"Meses: {months}\n"
                                         f"Temperatura: {temperatura}\n"
                                         f"Fecha de Medida: {resultados['fecha_medida']['dd/mm/yyyy']} ")
        root.destroy()

    root = Tk.Tk()
    root.title("Selector de Opciones")
    root.geometry("400x400")  # Tamaño inicial más pequeño para mostrar el scrollbar

    # Crear el Canvas
    canvas = Tk.Canvas(root)
    canvas.pack(side=Tk.LEFT, fill=Tk.BOTH, expand=True)

    # Crear la Scrollbar
    scrollbar = Tk.Scrollbar(root, orient=Tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=Tk.RIGHT, fill=Tk.Y)

    # Configurar el Canvas para usar la Scrollbar
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    # Crear un Frame dentro del Canvas para contener tu contenido
    frame = Tk.Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor=Tk.NW)

    # Agregar tu contenido al Frame
    # Botón para abrir el calendario
    Tk.Button(frame, text="Seleccionar fecha", command=abrir_calendario).pack(pady=5)

    # Label para mostrar la fecha seleccionada
    label_mostrar_fecha = Tk.Label(frame, text="Fecha seleccionada: ")
    label_mostrar_fecha.pack(pady=5)

    # Opciones de medida
    label_medida = Tk.Label(frame, text="Selecciona la medida:")
    label_medida.pack(pady=5)

    frame_medidas = Tk.Frame(frame)
    frame_medidas.pack()

    variable_medida = Tk.StringVar(value="Reflectancia")
    for medida in ["Reflectancia", "Absortancia", "Transmitancia CSP", "Transmitancia PV"]:
        Tk.Radiobutton(frame_medidas, text=medida, variable=variable_medida, value=medida).pack(side=Tk.LEFT)

    # Opciones de aparatos
    label_aparatos = Tk.Label(frame, text="Selecciona los aparatos:")
    label_aparatos.pack(pady=5)

    variables_aparatos = {aparato: Tk.BooleanVar() for aparato in ["FTIR", "Espectrofotómetro"]}
    variable_ventana = {aparato: Tk.StringVar(value="") for aparato in ["FTIR", "Espectrofotómetro"]}
    radio_ventana_si = {}
    radio_ventana_no = {}

    for aparato in ["FTIR", "Espectrofotómetro"]:
        Tk.Checkbutton(frame, text=aparato, variable=variables_aparatos[aparato],
                       command=lambda a=aparato: update_ventana_state(a)).pack(anchor='w')

        frame_ventana = Tk.Frame(frame)
        frame_ventana.pack(anchor='w', padx=20)

        radio_ventana_si[aparato] = Tk.Radiobutton(frame_ventana, text="Con ventana",
                                                   variable=variable_ventana[aparato], value="Con ventana",
                                                   state=Tk.DISABLED)
        radio_ventana_si[aparato].pack(side=Tk.LEFT)

        radio_ventana_no[aparato] = Tk.Radiobutton(frame_ventana, text="Sin ventana",
                                                   variable=variable_ventana[aparato], value="Sin ventana",
                                                   state=Tk.DISABLED)
        radio_ventana_no[aparato].pack(side=Tk.LEFT)

    # Selección de tipo de test
    label_test = Tk.Label(frame, text="Selecciona el tipo de test:")
    label_test.pack(pady=5)

    options_test = elegir_test_referencias("testsite")

    variable_test = Tk.StringVar(value="")
    menu_test = Tk.OptionMenu(frame, variable_test, *options_test)
    menu_test.pack(pady=5)

    # Selección de tipo de fabricante
    label_fabricante = Tk.Label(frame, text="Selecciona el fabricante:")
    label_fabricante.pack(pady=5)

    option_manufacturer = elegir_test_referencias("fabricantes")

    variable_fabricante = Tk.StringVar(value="")
    menu_test = Tk.OptionMenu(frame, variable_fabricante, *option_manufacturer)
    menu_test.pack(pady=5)

    # Selección de tipo de proyecto
    label_proyect = Tk.Label(frame, text="Selecciona el proyecto:")
    label_proyect.pack(pady=5)

    option_proyect = elegir_test_referencias("proyectos")

    variable_proyecto = Tk.StringVar(value="")
    menu_test = Tk.OptionMenu(frame, variable_proyecto, *option_proyect)
    menu_test.pack(pady=5)

    # Campo de entrada para las horas
    label_hours = Tk.Label(frame, text="Horas:")
    label_hours.pack(pady=5)

    entry_hours = Tk.Entry(frame)
    entry_hours.pack(pady=5)

    # Campo de entrada para meses
    label_months = Tk.Label(frame, text="Meses:")
    label_months.pack(pady=5)

    entry_months = Tk.Entry(frame)
    entry_months.pack(pady=5)

    # Campo de entrada para la temperatura
    label_t = Tk.Label(frame, text="Temperatura:")
    label_t.pack(pady=5)

    entry_t = Tk.Entry(frame)
    entry_t.pack(pady=5)

    # Botón para confirmar la selección
    button_submit = Tk.Button(frame, text="Guardar", command=submit)
    button_submit.pack(pady=20)

    # Ejecutar la aplicación
    root.mainloop()

    return resultados




def nombres_muestras_auto(paths_espectofotometro, path_ftir):
    nombres_muestras = []

    if paths_espectofotometro:
        for path in paths_espectofotometro:
            archivos = encontrar_archivos_asc_por_nombre("sample_", path)
            print("archivos encontrados: ", archivos)
            resultados = [re.search(r'sample_(.*)-[^-]+$', archivo).group(1) for archivo in archivos if
                          re.search(r'sample_(.*)-[^-]+\.Sample\.asc$', archivo)] #cambiar esto
            muestras_carpetas = list(set(resultados))
            nombres_muestras.extend(muestras_carpetas)
        #nombres_muestras_sin_duplicados = list(set(nombres_muestras))
        nombres_muestras_sin_duplicados = list(dict.fromkeys(nombres_muestras))
    else:
        print("buscando muestras en directorio ftir")
        # cadena de búsqueda para archivos y prefijo de pestañas
        cadena_busqueda_archivo = "data"
        prefijo_pestana = "sample_"

        # path_ftir puede ser una lista/iterable de rutas
        for directorio_ftir in path_ftir:
            print(f"Procesando directorio: {directorio_ftir}")
            directorio_path = Path(str(directorio_ftir))

            # buscar archivos Excel en subcarpetas (.xlsx y .xls)
            for ext in ("*.xlsx", "*.xls", ".xlsm"):
                for archivo_path in directorio_path.rglob(ext):
                    # comprobar que "data" esté en el nombre (case-insensitive)
                    if cadena_busqueda_archivo.lower() not in archivo_path.name.lower():
                        continue

                    ruta_completa_excel = str(archivo_path.resolve())
                    nombre_archivo = archivo_path.name
                    print("archivo muestras encontrado:", nombre_archivo)

                    try:
                        xls = pd.ExcelFile(ruta_completa_excel)
                    except Exception as e:
                        print(f"No se pudo leer {nombre_archivo}: {e}")
                        continue

                    # recorrer las pestañas y procesar las que empiecen por "sample_"
                    for pestana in xls.sheet_names:
                        if not isinstance(pestana, str):
                            continue
                        if not pestana.startswith(prefijo_pestana):
                            continue

                        cadena_a_procesar = pestana[len(prefijo_pestana):]  # quitar "sample_"


                        if "-" in cadena_a_procesar:
                            cadena_a_procesar = cadena_a_procesar.rsplit("-", 1)[0]

                        cadena_a_procesar = cadena_a_procesar.strip()
                        if cadena_a_procesar:
                            nombres_muestras.append(cadena_a_procesar)

        # obtener lista sin duplicados (opcional: ordenada)
        nombres_muestras_sin_duplicados = list(dict.fromkeys(nombres_muestras))  # mantiene orden de aparición

    return sorted(nombres_muestras_sin_duplicados)


def archivo_tfir(mensaje):
    # Seleccionar archivo Excel
    archivo_tfir = filedialog.askopenfilename(
        title=mensaje,
        filetypes=[("Archivos Excel", "*.xlsx;*.xls")]
    )

    if not archivo_tfir:
        print("No se seleccionó ningún archivo.")
        return None
    return archivo_tfir

def carpeta_con_archivos(mensaje):
    # Configurar Tkinter
    root = Tk.Tk()
    root.withdraw()  # Ocultar la ventana principal

    # Seleccionar carpeta
    carpeta_ppal = filedialog.askdirectory(
        title= mensaje,
        initialdir="."  # Iniciar en el directorio actual
    )

    # Si el usuario cancela la selección, filedialog.askdirectory() devuelve una cadena vacía
    if not carpeta_ppal:
        return [] # O podrías devolver None, o manejarlo como prefieras

    # Inicializar la lista de carpetas con la carpeta principal seleccionada
    carpetas = [carpeta_ppal]

    # Obtener todas las subcarpetas dentro de la carpeta seleccionada
    for sub in os.listdir(carpeta_ppal):
        full_path = os.path.join(carpeta_ppal, sub)
        if os.path.isdir(full_path):
            carpetas.append(full_path)

    # Devolver las rutas seleccionadas (incluyendo la principal si es que fue seleccionada)
    return carpetas

def encontrar_archivos_asc_por_nombre(nombre, carpeta):
    archivos_encontrados = []
    for raiz, directorios, archivos in os.walk(carpeta):
        for archivo in archivos:
            if archivo.startswith(nombre) and archivo.endswith(".asc"):
                ruta_completa = os.path.join(raiz, archivo)
                archivos_encontrados.append(ruta_completa)
    return archivos_encontrados

def espectro_medidas_zero_base_auto(muestra, carpetas, datos_basicos):
    print("Medidas espectofotómetro")
    for carpeta in carpetas:
        archivos_muestra = encontrar_archivos_asc_por_nombre("sample_"+ muestra + "-", carpeta)
        print("archivos muestra encontrados: ", archivos_muestra)
        if archivos_muestra and (("Reflectancia" in datos_basicos["medida"]) or ("Absortancia" in datos_basicos["medida"])):
            print("analizando carpeta: ", carpeta)
            archivos_base = encontrar_archivos_asc_por_nombre("base_", carpeta)
            archivos_zero = encontrar_archivos_asc_por_nombre("zero_", carpeta)
            archivos_ventana = encontrar_archivos_asc_por_nombre("ventana_", carpeta)
            #archivos_ventanabase = encontrar_archivos_asc_por_nombre("ventanabase_", carpeta)
            print("archivos base: ", archivos_base)
            print("archivos zero: ", archivos_zero)
            print("archivos ventana: ", archivos_ventana)
            if len(archivos_base) > 1:
                # Leer la fecha y hora de los archivos base y el archivo de muestra
                fechas_base = []
                for archivo in archivos_base:
                    with open(os.path.join(carpeta, archivo), 'r') as f:
                        lineas = f.readlines()
                        fecha_base = lineas[3].strip()  # Línea 3 (índice 2)
                        hora_base = lineas[4].strip()  # Línea 4 (índice 3)
                        datetime_base = datetime.strptime(f"{fecha_base} {hora_base}", "%y/%m/%d %H:%M:%S.%f")
                        fechas_base.append((archivo, datetime_base))
                print("fechas baseline: ", fechas_base)

                with open(os.path.join(carpeta, archivos_muestra[0]), 'r') as f:
                    lineas = f.readlines()
                    fecha_muestra = lineas[3].strip()  # Línea 3 (índice 2)
                    hora_muestra = lineas[4].strip()  # Línea 4 (índice 3)
                    datetime_muestra = datetime.strptime(f"{fecha_muestra} {hora_muestra}", "%y/%m/%d %H:%M:%S.%f")
                    print("fecha muestra: ", datetime_muestra)
                # Encontrar el archivo base con la fecha más cercana a la de la muestra
                archivo_base_cercano = min(fechas_base, key=lambda x: abs(x[1] - datetime_muestra))
                print("archivo base cercano: ", archivo_base_cercano)
                # Dejar solo el archivo base más cercano
                archivos_base = [archivo_base_cercano[0]]
                print("archivo base elegido: ", archivos_base)
            break
        else: # si es transmitancia, no se usan base y/o zero
            archivos_zero = ""
            archivos_base = ""
            archivos_ventana = ""
            #archivos_ventanabase = ""
    if archivos_muestra and (("Reflectancia" in datos_basicos["medida"]) or ("Absortancia" in datos_basicos["medida"])):
        archivos_zero_base = {"ZeroLine": archivos_zero[0], "BaseLine": archivos_base[0], "ventana":archivos_ventana}
    else:
        archivos_zero_base = {"ZeroLine": archivos_zero, "BaseLine": archivos_base, "ventana": archivos_ventana}

    return archivos_zero_base, archivos_muestra

def buscar_hojas(hojas, patrones):
    return {patron: [hoja for hoja in hojas if hoja.startswith(patron)] for patron in patrones}

def seleccionar_unico_valor(diccionario, claves):
    for clave in claves:
        valor = diccionario.get(clave)
        diccionario[clave] = valor[0] if isinstance(valor, list) and valor else None


def procesar_cadena(cadena):
    """
    Convierte una cadena a minúsculas y elimina todos los caracteres separadores y espacios.

    Args:
      cadena: La cadena de texto a procesar.

    Returns:
      La cadena procesada en minúsculas y sin separadores ni espacios.
    """
    cadena = cadena.lower()
    cadena = cadena.strip()
    separadores = " ,.-_/"  # Define los separadores a eliminar
    for separador in separadores:
        cadena = cadena.replace(separador, "")  # Reemplaza con una cadena vacía (elimina)
    return cadena

def buscar_y_filtrar_excel_data_ftir(directorios_ftir, nombre_muestra):
    """
    Busca archivos Excel con "data" en su nombre en una lista de directorios
    y sus subcarpetas. Luego, en cada uno de esos archivos, identifica las
    pestañas que contienen "sample_" + nombre_muestra.

    Args:
        directorios_ftir (list): Una lista de rutas a los directorios donde buscar.
        nombre_muestra (str): La cadena que representa el nombre de la muestra a buscar
                              en los nombres de las pestañas.

    Returns:
        tuple: Una tupla que contiene:
               - dict: Un diccionario donde las claves son las rutas completas de los archivos Excel
                       encontrados, y los valores son listas de nombres de pestañas que cumplen el criterio.
               - str or None: La fecha extraída del primer archivo encontrado como 'YYYYMMDD' string.
                              Será None si no se encuentra ningún archivo relevante con fecha.
    """
    print("Búsqueda de muestras IR")
    archivos_excel_filtrados = {}
    fecha_medida = None
    #nombre_muestra = procesar_cadena(nombre_muestra) no hace falta, ya está procesado
    cadena_busqueda_archivo = "data"  # Se busca "data" en lugar de "_data_" según la corrección.
    cadena_busqueda_pestana = "sample_" + nombre_muestra

    for directorio_ftir in directorios_ftir:
        print(f"Procesando directorio: {directorio_ftir}")  # Para debug
        directorio_path = Path(str(directorio_ftir))

        # 2. Barrido de archivos en el directorio actual y subcarpetas.
        for archivo_path in directorio_path.rglob("*.xlsx"):
            if cadena_busqueda_archivo in archivo_path.name:
                ruta_completa_excel = str(archivo_path.resolve())
                nombre_archivo = archivo_path.name
                print("archivo muestras encontrado: ", nombre_archivo)
                try:
                    xls = pd.ExcelFile(ruta_completa_excel)
                    nombres_pestanas = xls.sheet_names
                    pestanas_encontradas = []
                    for pestana in nombres_pestanas:
                        cadena_a_procesar = pestana

                        # 1. Remove "sample_" if present
                        if cadena_a_procesar.startswith("sample_"):
                            cadena_a_procesar = cadena_a_procesar[len("sample_"):]

                        # 2. Remove everything after the last hyphen if present
                        if "-" in cadena_a_procesar:
                            cadena_a_procesar = cadena_a_procesar.rsplit('-', 1)[0]

                        # 3. Process the modified string
                        pestana_procesado = procesar_cadena(cadena_a_procesar)
                        pestana_procesado = "sample_"+pestana_procesado
                        print("Buscando pestana: ", cadena_busqueda_pestana)
                        print("pestana: ", pestana, " y nombre procesado: ", pestana_procesado)
                        pattern = re.compile(r"^{}(_.*|$)".format(re.escape(cadena_busqueda_pestana)))
                        if pattern.match(pestana_procesado):
                            pestanas_encontradas.append(pestana)
                    if pestanas_encontradas:
                        archivos_excel_filtrados[ruta_completa_excel] = pestanas_encontradas
                        regex_data = r"^(.*?)data"
                        coincidencia = re.search(regex_data, nombre_archivo)
                        # Intenta extraer la fecha del nombre del archivo
                        if coincidencia:
                            fecha_medida = coincidencia.group(1)
                        else:
                            print(f"No se encontró el patrón 'data' en el nombre del archivo: {nombre_archivo}")
                        break  # Importante: Termina el bucle for actual si se encuentran pestañas
                    else:
                         print(f"No se encontraron pestañas que coincidan con '{cadena_busqueda_pestana}' en el archivo: {nombre_archivo}")
                except Exception as e:
                    print(f"Error al procesar el archivo Excel '{ruta_completa_excel}': {e}")
                    continue  # Continúa con el siguiente archivo
            # Este else se ejecuta si el bucle for interno *no* se completó normalmente
            else:
                print(f"No se encontró ningún archivo que coincida en el directorio: {directorio_ftir}")

    return archivos_excel_filtrados, fecha_medida

def buscar_archivos_ref_y_pestanas(directorios_base, lista_patrones, fecha_medida, path_medidas_ftir_muestras):
    """
    Busca archivos Excel que contengan "_ref_" en su nombre y la fecha de medida.
    Primero en los directorios base, y si no se encuentra, usa path_medidas_ftir_muestras.
    Luego identifica las pestañas que coinciden con cualquiera de los patrones.
    """
    archivos_ref_filtrados = {}
    cadena_busqueda_archivo_ref = fecha_medida + "ref"
    print("Búsqueda de referencias IR")
    print("Buscando archivos que contengan:", cadena_busqueda_archivo_ref)

    # Función auxiliar para buscar archivos en un directorio
    def buscar_en_directorio(directorio):
        resultados = {}
        directorio_path = Path(directorio)
        for archivo_path in directorio_path.rglob("*.xlsx"):
            if cadena_busqueda_archivo_ref in archivo_path.name:
                ruta_completa_excel_ref = str(archivo_path.resolve())
                try:
                    xls = pd.ExcelFile(ruta_completa_excel_ref)
                    nombres_pestanas = xls.sheet_names
                    pestanas_encontradas_ref = [
                        pestana for pestana in nombres_pestanas if any(patron in pestana for patron in lista_patrones)
                    ]
                    if pestanas_encontradas_ref:
                        resultados[ruta_completa_excel_ref] = pestanas_encontradas_ref
                        break
                except Exception as e:
                    print(f"Error al procesar el archivo Excel '{ruta_completa_excel_ref}': {e}")
                    continue
        return resultados

    # 1. Buscar en los directorios base
    for directorio_base in directorios_base:
        print(f"Procesando directorio base: {directorio_base}")
        archivos_ref_filtrados.update(buscar_en_directorio(directorio_base))

    # 2. Si no se encontró ningún archivo ref, usar el archivo de path_medidas_ftir_muestras
    if not archivos_ref_filtrados:
        print("No se encontró ningún archivo ref en los directorios base.")
        ruta_archivo = list(path_medidas_ftir_muestras.keys())[0]
        print(f"Usando archivo de path_medidas_ftir_muestras: {ruta_archivo}")
        try:
            xls = pd.ExcelFile(ruta_archivo)
            nombres_pestanas = xls.sheet_names
            pestanas_encontradas_ref = [
                pestana for pestana in nombres_pestanas if any(patron in pestana for patron in lista_patrones)
            ]
            if pestanas_encontradas_ref:
                archivos_ref_filtrados[ruta_archivo] = pestanas_encontradas_ref
        except Exception as e:
            print(f"Error al procesar el archivo Excel '{ruta_archivo}': {e}")

    return archivos_ref_filtrados

def ftir_medidas_auto(path_ftir, nombre_muestra, ventana):

    path_medidas_ftir_muestras, fecha_medida = buscar_y_filtrar_excel_data_ftir(path_ftir, nombre_muestra)
    print("path medidas: ")
    print(path_medidas_ftir_muestras)
    "path_medidas_ftir_muestras = diccionario {ruta: pestana}"
    print("fecha medida: ", fecha_medida)
    if path_medidas_ftir_muestras:
        #patrones_comunes = ["zero_", "baseoro_"]
        #patrones_ventana = ["basenegro_", "ventana_", "ventanaoro_", "ventananegro_"]
        patrones_comunes = ["zero_", "base_"]
        patrones_ventana = ["ventana_"]

        patrones = patrones_comunes
        if ventana:
            patrones += patrones_ventana

        archivos_ref_ftir = buscar_archivos_ref_y_pestanas(path_ftir, patrones, fecha_medida, path_medidas_ftir_muestras)
        print("archivos_ref_ftir: ", archivos_ref_ftir)
        archivos_ir = { #cambiar, que sea solo base_
            "path": path_ftir,
            "zero": {},
            "base": {},
            "ventana": {},
            "path_muestras": path_medidas_ftir_muestras,
            "path_referencias": archivos_ref_ftir,
        }

        for ruta_archivo, pestañas in archivos_ref_ftir.items():
            for pestana in pestañas:
                if pestana.startswith("zero_"):
                    if "zero" not in archivos_ir:
                        archivos_ir["zero"] = {}
                    archivos_ir["zero"][ruta_archivo] = pestana
                elif pestana.startswith("base_"):
                    if "base" not in archivos_ir:
                        archivos_ir["base"] = {}
                    archivos_ir["base"][ruta_archivo] = pestana
                elif pestana.startswith("ventana_"):
                    if "ventana" not in archivos_ir:
                        archivos_ir["ventana"] = {}
                    archivos_ir["ventana"][ruta_archivo] = pestana

        print("archivos ir: ", archivos_ir)
    else:
        archivos_ir = {}
        print("archivos ir: ",archivos_ir)
    return archivos_ir


def elegir_columnas_referencia(nombre_pestana, mensaje):
    archivo_referencias = leer_archivo_referencias()

    try:
        data_ref = pd.read_excel(archivo_referencias, sheet_name=nombre_pestana)
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo references.xlsx")
        return None
    except ValueError:
        print(f"Error: No se encontró la pestaña '{nombre_pestana}' en references.xlsx")
        return None
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return None

    columnas = data_ref.columns.tolist()

    if not columnas:
        print(f"Error: La pestaña '{nombre_pestana}' no tiene columnas.")
        return None

    ventana = Tk.Tk()
    ventana.title(mensaje)

    nombre_elegido = Tk.StringVar()  # Variable para almacenar la selección

    label_columna = Tk.Label(ventana, text=mensaje)
    label_columna.pack(pady=5)

    combo_columna = ttk.Combobox(ventana, values=columnas, state="readonly")
    combo_columna.pack(pady=5)
    combo_columna.current(0)  # Selecciona por defecto el primer elemento

    def submit():
        nombre_elegido.set(combo_columna.get())  # Asigna el valor seleccionado
        ventana.quit()  # Cierra el loop de Tkinter
        ventana.destroy()  # Cierra la ventana

    boton_seleccionar = Tk.Button(ventana, text="Seleccionar", command=submit)
    boton_seleccionar.pack(pady=10)

    ventana.mainloop()  # Inicia la interfaz gráfica

    return nombre_elegido.get()  # Devuelve la columna seleccionada


def elegir_test_referencias(nombre_pestana):
    archivo_referencias = leer_archivo_referencias()
    try:
        data_ref = pd.read_excel(archivo_referencias, sheet_name=nombre_pestana, header=None) # Lee sin encabezados
        if data_ref.empty:
            print(f"Advertencia: La pestaña '{nombre_pestana}' está vacía.")
            return []  # Devuelve una lista vacía si la pestaña está vacía
        else:
            return data_ref.iloc[:, 0].tolist()  # Devuelve la primera columna como lista
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo references.xlsx")
        return None
    except ValueError:
        print(f"Error: No se encontró la pestaña '{nombre_pestana}' en references.xlsx")
        return None
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return None


def preguntar_dato_simple(titulo, pregunta):
    # Crear la ventana principal
    root = Tk.Tk()
    root.withdraw()  # Ocultar la ventana principal

    # Obtener el nombre de la muestra desde un cuadro de diálogo
    dato = simpledialog.askstring(title=titulo, prompt=pregunta)
    return dato



#dependiendo de si es UV o IR, se lee una wavelength u otra
    #guardarlo en muestra self.spectra_ref_UV, self.spectra_ref_IR
    #opción de que no haya infrarrojo
    #resto de programa: si no hay infrarrojo, se salta toda la parte de IR en muestas.