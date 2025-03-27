from tkinter import simpledialog
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import tkinter as Tk
from tkinter import messagebox
from tkinter import ttk
from tkcalendar import Calendar


def leer_archivo_referencias():
    directorio_script = os.path.dirname(os.path.realpath(__file__))
    archivo_referencias = os.path.normpath(
        os.path.join(directorio_script, "../para_el_usuario/references.xlsx"))

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
    archivos_ventana = encontrar_archivos_asc_por_nombre("ventana_", carpeta)
    archivos_ventanabase = encontrar_archivos_asc_por_nombre("ventanabase_", carpeta)
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

    archivos_zero_base = {"ZeroLine": archivos_zero[0], "BaseLine": archivos_base[0], "ventana":archivos_ventana,"ventanabase":archivos_ventanabase}
    
    return archivos_zero_base, archivos_muestra

def buscar_hojas(hojas, patrones):
    return {patron: [hoja for hoja in hojas if hoja.startswith(patron)] for patron in patrones}

def seleccionar_unico_valor(diccionario, claves):
    for clave in claves:
        valor = diccionario.get(clave)
        diccionario[clave] = valor[0] if isinstance(valor, list) and valor else None

def ftir_medidas_auto(archivos_ir, muestra, archivo_ftir, ventana, tipo_ftir):
    # Leer las hojas del archivo Excel
    try:
        hojas = pd.ExcelFile(archivo_ftir).sheet_names
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo Excel.\n{e}")
        return None

    patrones_comunes = ["zero_", "baseoro_"]
    patrones_ventana = ["basenegro_", "ventana_", "ventanaoro_", "ventananegro_"]

    if tipo_ftir == "muestras":
        hojas_encontradas = {"muestras": [hoja for hoja in hojas if hoja.startswith(muestra)]}
        print("hojas encontradas:", hojas_encontradas)

    elif tipo_ftir == "referencias":
        patrones = patrones_comunes
        if ventana:
            patrones += patrones_ventana

        hojas_encontradas = buscar_hojas(hojas, patrones)
        seleccionar_unico_valor(hojas_encontradas, ["baseoro_", "zero_", "basenegro_", "ventana_", "ventanaoro_", "ventananegro_"])

    archivos_ir.update(hojas_encontradas)

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