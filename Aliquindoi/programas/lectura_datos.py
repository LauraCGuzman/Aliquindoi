from tkinter import simpledialog
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import tkinter as Tk
from tkinter import messagebox
from tkinter import ttk
from tkcalendar import Calendar

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
        resultados["hours"] = hours
        resultados["months"] = months
        resultados["temperatura"] = float(temperatura)

        # Obtener la fecha seleccionada
        fecha_seleccionada = cal.get_date()
        fecha_dt = datetime.strptime(fecha_seleccionada, "%d/%m/%Y")

        # Formatear la fecha en los formatos requeridos
        fecha_formato1 = fecha_dt.strftime("%d/%m/%Y")
        fecha_formato2 = fecha_dt.strftime("%Y%m%d")

        resultados["fecha_medida"] = {"dd/mm/yyyy": fecha_formato1, "yyyyMMdd": fecha_formato2}

        # Obtener selección de aparatos y si tienen ventana o no
        for aparato in ["FTIR", "Espectrofotómetro"]:
            if variables_aparatos[aparato].get():
                resultados["aparatos"][aparato] = variable_ventana[aparato].get()

        if not resultados["aparatos"]:
            messagebox.showwarning("Advertencia", "Debes seleccionar al menos un aparato.")
            return

        messagebox.showinfo("Selección", f"Has seleccionado:\nMedida: {selected_medida}\n"
                                         f"Aparatos: {resultados['aparatos']}\n"
                                         f"Tipo de test: {selected_test}\n"
                                         f"Fabricante: {selected_fabricante}\n"
                                         f"Proyecto: {selected_proyecto}\n"
                                         f"Horas: {hours}\n"
                                         f"Meses: {months}\n"
                                         f"Temperatura: {temperatura}\n"
                                         f"Fecha de Medida: {fecha_formato1} ")
        root.destroy()

    # Crear ventana principal
    root = Tk.Tk()
    root.title("Selector de Opciones")
    root.geometry("400x700")

    # Añadir el selector de fecha
    label_fecha = Tk.Label(root, text="Fecha de Medida:")
    label_fecha.pack(pady=5)

    cal = Calendar(root, selectmode="day", date_pattern="dd/mm/yyyy")
    cal.pack(pady=5)

    # Opciones de medida
    label_medida = Tk.Label(root, text="Selecciona la medida:")
    label_medida.pack(pady=5)

    variable_medida = Tk.StringVar(value="Reflectancia")
    for medida in ["Reflectancia", "Absortancia", "Transmitancia"]:
        Tk.Radiobutton(root, text=medida, variable=variable_medida, value=medida).pack(anchor='w')

    # Opciones de aparatos
    label_aparatos = Tk.Label(root, text="Selecciona los aparatos:")
    label_aparatos.pack(pady=5)

    variables_aparatos = {aparato: Tk.BooleanVar() for aparato in ["FTIR", "Espectrofotómetro"]}
    variable_ventana = {aparato: Tk.StringVar(value="") for aparato in ["FTIR", "Espectrofotómetro"]}
    radio_ventana_si = {}
    radio_ventana_no = {}

    for aparato in ["FTIR", "Espectrofotómetro"]:
        Tk.Checkbutton(root, text=aparato, variable=variables_aparatos[aparato],
                       command=lambda a=aparato: update_ventana_state(a)).pack(anchor='w')

        frame_ventana = Tk.Frame(root)
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
    label_test = Tk.Label(root, text="Selecciona el tipo de test:")
    label_test.pack(pady=5)

    options_test = elegir_test_referencias("testsite")

    variable_test = Tk.StringVar(value="")
    menu_test = Tk.OptionMenu(root, variable_test, *options_test)
    menu_test.pack(pady=5)

    # Selección de tipo de fabricante
    label_fabricante = Tk.Label(root, text="Selecciona el fabricante:")
    label_fabricante.pack(pady=5)

    option_manufacturer = elegir_test_referencias("fabricantes")

    variable_fabricante = Tk.StringVar(value="")
    menu_test = Tk.OptionMenu(root, variable_fabricante, *option_manufacturer)
    menu_test.pack(pady=5)

    # Selección de tipo de proyecto
    label_proyect = Tk.Label(root, text="Selecciona el proyecto:")
    label_proyect.pack(pady=5)

    option_proyect = elegir_test_referencias("proyectos")

    variable_proyecto = Tk.StringVar(value="")
    menu_test = Tk.OptionMenu(root, variable_proyecto, *option_proyect)
    menu_test.pack(pady=5)

    # Campo de entrada para las horas
    label_hours = Tk.Label(root, text="Horas:")
    label_hours.pack(pady=5)

    entry_hours = Tk.Entry(root)
    entry_hours.pack(pady=5)

    # Campo de entrada para meses
    label_months = Tk.Label(root, text="Meses:")
    label_months.pack(pady=5)

    entry_months = Tk.Entry(root)
    entry_months.pack(pady=5)

    # Campo de entrada para la temperatura
    label_t = Tk.Label(root, text="Temperatura:")
    label_t.pack(pady=5)

    entry_t = Tk.Entry(root)
    entry_t.pack(pady=5)

    # Botón para confirmar la selección
    button_submit = Tk.Button(root, text="Guardar", command=submit)
    button_submit.pack(pady=20)

    # Ejecutar la aplicación
    root.mainloop()

    return resultados


def detener_analisis():
    respuesta = preguntar_dato_simple("Análisis de muestras", "Desea seguir analizando muestras? (y/n): ")

    if (respuesta == "n") | (respuesta == 0) | (respuesta == "no"):
        respuesta = False
        print("Fin del programa")
    return respuesta

def nombres_muestras_auto():
    """
        Cuadro de diálogo para ingresar múltiples nombres de muestras.

        Returns:
            list: Lista de nombres de muestras ingresados.
        """

    ventana = Tk.Tk()
    ventana.title("Ingreso de Nombres de Muestras")

    nombres = []

    def agregar_nombre():
        nombre = entrada.get()
        if nombre:
            nombres.append(nombre)
            lista_nombres.insert(Tk.END, nombre)
            entrada.delete(0, Tk.END)

    def borrar_nombre():
        seleccion = lista_nombres.curselection()
        if seleccion:
            index = seleccion[0]
            lista_nombres.delete(index)
            nombres.pop(index)

    def guardar_nombres():
        ventana.destroy()  # Cierra la ventana

    # Label para indicar qué ingresar
    etiqueta = Tk.Label(ventana, text="Ingrese el nombre de la muestra:")
    etiqueta.pack()

    # Campo de entrada
    entrada = Tk.Entry(ventana)
    entrada.pack()

    # Botón para agregar nombre
    boton_agregar = Tk.Button(ventana, text="+", command=agregar_nombre)
    boton_agregar.pack()

    # Lista para mostrar los nombres ingresados
    lista_nombres = Tk.Listbox(ventana)
    lista_nombres.pack()

    # Botón para borrar nombre seleccionado
    boton_borrar = Tk.Button(ventana, text="Borrar", command=borrar_nombre)
    boton_borrar.pack()

    # Botón para guardar los nombres
    boton_guardar = Tk.Button(ventana, text="Guardar", command=guardar_nombres)
    boton_guardar.pack()

    ventana.mainloop()

    return nombres

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

def carpeta_espectofotometro():
    # Configurar Tkinter
    root = Tk.Tk()
    root.withdraw()  # Ocultar la ventana principal

    # Seleccionar carpeta
    carpeta = filedialog.askdirectory(
        title="Selecciona la carpeta de datos del espectrofotómetro",
        initialdir="."  # Iniciar en el directorio actual
    )

    # Devolver las rutas seleccionadas
    return carpeta

def encontrar_archivos_asc_por_nombre(nombre, carpeta):
    archivos_encontrados = []
    for raiz, directorios, archivos in os.walk(carpeta):
        for archivo in archivos:
            if archivo.startswith(nombre) and archivo.endswith(".asc"):
                ruta_completa = os.path.join(raiz, archivo)
                archivos_encontrados.append(ruta_completa)
    return archivos_encontrados

def espectro_medidas_zero_base_auto(muestra, carpeta):

    archivos_base = encontrar_archivos_asc_por_nombre("base_", carpeta)
    archivos_zero = encontrar_archivos_asc_por_nombre("zero_", carpeta)
    archivos_muestra = encontrar_archivos_asc_por_nombre(muestra, carpeta)

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

        with open(os.path.join(carpeta, archivos_muestra[0]), 'r') as f:
            lineas = f.readlines()
            fecha_muestra = lineas[3].strip()  # Línea 3 (índice 2)
            hora_muestra = lineas[4].strip()  # Línea 4 (índice 3)
            datetime_muestra = datetime.strptime(f"{fecha_muestra} {hora_muestra}", "%y/%m/%d %H:%M:%S.%f")

        # Encontrar el archivo base con la fecha más cercana a la de la muestra
        archivo_base_cercano = min(fechas_base, key=lambda x: abs(x[1] - datetime_muestra))

        # Dejar solo el archivo base más cercano
        archivos_base = [archivo_base_cercano[0]]

    archivos_zero_base = {"ZeroLine": archivos_zero[0], "BaseLine": archivos_base[0]}
    
    return archivos_zero_base, archivos_muestra

def buscar_hojas(hojas, patrones):
    return {patron: [hoja for hoja in hojas if hoja.startswith(patron)] for patron in patrones}

def seleccionar_unico_valor(diccionario, claves):
    for clave in claves:
        diccionario[clave] = diccionario[clave][0] if diccionario[clave] else None

def ftir_medidas_auto(archivos_ir, muestra, archivo_tfir, ventana=False, tipo_ftir="ambos"):
    # Leer las hojas del archivo Excel
    try:
        hojas = pd.ExcelFile(archivo_tfir).sheet_names
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo Excel.\n{e}")
        return None

    patrones_comunes = ["zero_"]
    patrones_ventana = ["refnegro_", "ventana_", "ventanaoro_", "ventananegro_"]

    if tipo_ftir == "ambos":
        patrones = ["reforo_", muestra] + patrones_comunes
        if ventana:
            patrones += patrones_ventana
    elif tipo_ftir == "muestras":
        patrones = [muestra]
    elif tipo_ftir == "referencias":
        patrones = ["base_"] + patrones_comunes
        if ventana:
            patrones += patrones_ventana

    hojas_encontradas = buscar_hojas(hojas, patrones)
    seleccionar_unico_valor(hojas_encontradas, ["reforo_", "zero_", "base_", "refnegro_", "ventana_", "ventanaoro_", "ventananegro_"])

    archivos_ir.update(hojas_encontradas)

    return archivos_ir

def elegir_columnas_referencia(nombre_pestana, mensaje):
    try:
        data_ref = pd.read_excel("references.xlsx", sheet_name=nombre_pestana)
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

    columna_elegida = Tk.StringVar()  # Variable para almacenar la selección

    def seleccionar_columna():
        ventana.quit()  # Finaliza el bucle de eventos sin cerrar la ventana

    # Crear la ventana
    ventana = Tk.Tk()
    ventana.title(mensaje)

    label_columna = Tk.Label(ventana, text="Seleccione la columna:")
    label_columna.pack(pady=5)

    combo_columna = ttk.Combobox(ventana, values=columnas, state="readonly", textvariable=columna_elegida)
    combo_columna.current(0)
    combo_columna.pack(pady=5)

    boton_seleccionar = Tk.Button(ventana, text="Seleccionar", command=seleccionar_columna)
    boton_seleccionar.pack(pady=10)

    ventana.mainloop()  # Inicia la interfaz gráfica

    ventana.destroy()  # Cierra la ventana completamente

    return columna_elegida.get()  # Devuelve la columna seleccionada

def elegir_test_referencias(nombre_pestana):
    try:
        data_ref = pd.read_excel("references.xlsx", sheet_name=nombre_pestana, header=None) # Lee sin encabezados
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


#funciones no usadas, pendientes de revisar
def preguntar_dato_simple(titulo, pregunta):
    # Crear la ventana principal
    root = Tk.Tk()
    root.withdraw()  # Ocultar la ventana principal

    # Obtener el nombre de la muestra desde un cuadro de diálogo
    dato = simpledialog.askstring(title=titulo, prompt=pregunta)
    return dato
def nombre_muestra():
    '''
      Programa que pide al usuario mediante una ventana emergente el nombre de la muestra
      y la guarda en la instancia "muestra".
    :return:
    '''
    nombre_muestra = preguntar_dato_simple("Nombre de la muestra", "Introduzca el nombre de la muestra:")

    return nombre_muestra
def archivo_tfir_base_zero_samples():
    # Seleccionar archivo Excel
    archivo_tfir = filedialog.askopenfilename(
        title="Seleccionar archivo Excel ftir",
        filetypes=[("Archivos Excel", "*.xlsx;*.xls")]
    )

    if not archivo_tfir:
        print("No se seleccionó ningún archivo.")
        return None

    # Leer las hojas del archivo Excel
    try:
        hojas = pd.ExcelFile(archivo_tfir).sheet_names
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo Excel.\n{e}")
        return None

    # Crear la ventana de selección de hojas
    ventana_seleccion = Tk.Tk()
    ventana_seleccion.title("Seleccionar hojas")
    ventana_seleccion.geometry("400x500")  # Ajustar tamaño de la ventana

    # Variables para las selecciones
    baseline_var = Tk.StringVar(value=hojas[0])  # Seleccionar la primera hoja por defecto
    zeroline_var = Tk.StringVar(value=hojas[0])

    # Etiqueta y desplegable para baseline
    ttk.Label(ventana_seleccion, text="Seleccionar hoja Baseline:").pack(pady=5)
    baseline_menu = ttk.OptionMenu(ventana_seleccion, baseline_var, None, *hojas)
    baseline_menu.pack(pady=5)

    # Etiqueta y desplegable para zeroline
    ttk.Label(ventana_seleccion, text="Seleccionar hoja Zeroline:").pack(pady=5)
    zeroline_menu = ttk.OptionMenu(ventana_seleccion, zeroline_var, None, *hojas)
    zeroline_menu.pack(pady=5)

    # Etiqueta y lista de selección múltiple para muestras
    ttk.Label(ventana_seleccion, text="Seleccionar hojas Muestras:").pack(pady=5)
    listbox_muestras = Tk.Listbox(ventana_seleccion, selectmode=Tk.MULTIPLE, height=15, exportselection=False)
    listbox_muestras.pack(pady=5, fill=Tk.BOTH, expand=True)

    # Cargar las hojas en el Listbox
    for hoja in hojas:
        listbox_muestras.insert(Tk.END, hoja)

    # Variable para almacenar el resultado
    resultado = {"archivo": None, "baseline": None, "zeroline": None, "muestras": []}

    def confirmar_seleccion():
        # Obtener las selecciones
        baseline = baseline_var.get()
        zeroline = zeroline_var.get()
        muestras_indices = listbox_muestras.curselection()
        if muestras_indices:
            muestras = [hojas[i] for i in muestras_indices]
        else:
            muestras = []

        # Validar las selecciones
        if not baseline:
            messagebox.showerror("Error", "Debe seleccionar una hoja como Baseline.")
            return
        if not zeroline:
            messagebox.showerror("Error", "Debe seleccionar una hoja como Zeroline.")
            return
        if not muestras:
            messagebox.showerror("Error", "Debe seleccionar al menos una hoja como Muestra.")
            return

        # Guardar las selecciones en el resultado
        resultado["archivo"] = archivo_tfir
        resultado["baseline"] = baseline
        resultado["zeroline"] = zeroline
        resultado["muestras"] = muestras

        # Cerrar la ventana
        ventana_seleccion.quit()

    # Botón para confirmar las selecciones
    ttk.Button(ventana_seleccion, text="Confirmar selección", command=confirmar_seleccion).pack(pady=10)

    ventana_seleccion.mainloop()
    ventana_seleccion.destroy()  # Asegurar que la ventana se destruya completamente

    # Devolver el resultado
    return resultado


def archivos_espectofotometro():
    print("leyendo resultados del espectofotómetro")
    # Configurar Tkinter
    root = Tk.Tk()
    root.withdraw()  # Ocultar la ventana principal

    # Seleccionar archivo para ZeroLine y BaseLine
    archivos_zero_base = {}
    for tipo in ['ZeroLine', 'BaseLine']:
        print(f"Seleccione el archivo para {tipo}")
        archivos_zero_base[tipo] = filedialog.askopenfilename(
            title=f"Selecciona el archivo para {tipo}",
            filetypes=[("Archivos ASC", "*.asc")]
        )

    # Seleccionar múltiples archivos para las medidas
    file_paths_muestras = filedialog.askopenfilenames(
        title="Selecciona los archivos para todas las medidas",
        filetypes=[("Archivos ASC", "*.asc")]
    )

    # Devolver las rutas seleccionadas
    return archivos_zero_base, file_paths_muestras

def seleccionar_spectro_referencia_IR_UV():
    # Definir las listas de columnas
    columns_IR = ["FS1 (specular FISE absorber)", "FS2 (diffuse BSII absorber)", "Gold specular",
                  "Gold diffuse RT-Au02c", "None"]
    columns_UV = ["Grey spectralon 10% (provided by Octalab in 2021)", "ref2"]


    # Crear la ventana de selección de hojas
    ventana_seleccion = Tk.Tk()
    ventana_seleccion.title("Seleccionar espectros de referencia")
    ventana_seleccion.geometry("400x400")  # Ajustar tamaño de la ventana

    # Variables para las selecciones
    ir_var = Tk.StringVar(value= columns_IR)  # Seleccionar la primera hoja por defecto
    uv_var = Tk.StringVar(value= columns_UV)

    # Etiqueta y desplegable para baseline
    ttk.Label(ventana_seleccion, text="Seleccionar referencia IR").pack(pady=5)
    baseline_menu = ttk.OptionMenu(ventana_seleccion, ir_var, None, *columns_IR)
    baseline_menu.pack(pady=5)

    # Etiqueta y desplegable para zeroline
    ttk.Label(ventana_seleccion, text="Seleccionar referencia UV:").pack(pady=5)
    zeroline_menu = ttk.OptionMenu(ventana_seleccion, uv_var, None, *columns_UV)
    zeroline_menu.pack(pady=5)

    # Variable para almacenar el resultado
    resultado = {"ir_var": ir_var, "uv_var": uv_var}

    def confirmar_seleccion():
        # Obtener las selecciones
        ir_selection = ir_var.get()
        uv_selection = uv_var.get()

        # Validar las selecciones
        if not ir_selection:
            messagebox.showerror("Error", "Debe seleccionar una referencia IR.")
            return
        if not uv_selection:
            messagebox.showerror("Error", "Debe seleccionar una referencia UV.")
            return

        # Guardar las selecciones en el resultado
        resultado["ir_var"] = ir_selection
        resultado["uv_var"] = uv_selection

        # Cerrar la ventana
        ventana_seleccion.quit()

    # Botón para confirmar las selecciones
    ttk.Button(ventana_seleccion, text="Confirmar selección", command=confirmar_seleccion).pack(pady=10)

    ventana_seleccion.mainloop()
    ventana_seleccion.destroy()  # Asegurar que la ventana se destruya completamente

    # Devolver el resultado
    return resultado
def seleccionar_spectro_referencia_UV():
    # Definir las listas de columnas
    columns_UV = ["Grey spectralon 10% (provided by Octalab in 2021)", "ref2"]


    # Crear la ventana de selección de hojas
    ventana_seleccion = Tk.Tk()
    ventana_seleccion.title("Seleccionar espectros de referencia")
    ventana_seleccion.geometry("400x400")  # Ajustar tamaño de la ventana

    # Variables para las selecciones
    uv_var = Tk.StringVar(value= columns_UV)

    # Etiqueta y desplegable para zeroline
    ttk.Label(ventana_seleccion, text="Seleccionar referencia UV:").pack(pady=5)
    zeroline_menu = ttk.OptionMenu(ventana_seleccion, uv_var, None, *columns_UV)
    zeroline_menu.pack(pady=5)

    # Variable para almacenar el resultado
    resultado = {"ir_var": "NaN", "uv_var": uv_var}

    def confirmar_seleccion():
        # Obtener las selecciones
        uv_selection = uv_var.get()

        # Validar las selecciones
        if not uv_selection:
            messagebox.showerror("Error", "Debe seleccionar una referencia UV.")
            return

        # Guardar las selecciones en el resultado
        resultado["ir_var"] = "NaN"
        resultado["uv_var"] = uv_selection

        # Cerrar la ventana
        ventana_seleccion.quit()

    # Botón para confirmar las selecciones
    ttk.Button(ventana_seleccion, text="Confirmar selección", command=confirmar_seleccion).pack(pady=10)

    ventana_seleccion.mainloop()
    ventana_seleccion.destroy()  # Asegurar que la ventana se destruya completamente

    # Devolver el resultado
    return resultado



#dependiendo de si es UV o IR, se lee una wavelength u otra
    #guardarlo en muestra self.spectra_ref_UV, self.spectra_ref_IR
    #opción de que no haya infrarrojo
    #resto de programa: si no hay infrarrojo, se salta toda la parte de IR en muestas.