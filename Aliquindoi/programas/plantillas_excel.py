import os
import xlwings as xw

def leer_asc_para_exportar_excel(path_asc):
    with open(path_asc, 'r') as file:
        lines = file.readlines()

    data_start = False
    data = []

    for i, line in enumerate(lines):
        if "#DATA" in line:
            data_start = True
        if not data_start:
            data.append(line.strip().replace(",", "."))  # Guardar la línea tal cual antes de #DATA
        else:
            split_line = line.split()
            if len(split_line) > 1:  # Asegurar que hay al menos dos elementos
                cleaned_value = split_line[1].replace(",", ".")  # Reemplazar , por .
                data.append(cleaned_value)  # Guardar la segunda columna ya corregida

    # **Eliminar las filas 74 a 79 si hay suficientes datos**
    del data[73:79]  # Elimina de la posición 74 a 79 inclusive
    del data[86]
    data[86:86] = ["", ""]  # Inserta dos elementos vacíos en la posición 86
    
    return data

def elegir_plantilla(Muestra):
    path = Muestra.path_zero
    with open(path, 'r') as file:
        lines = file.readlines()

    linea = lines[-1]
    split_line = linea.split()
    medida = split_line[0].replace(",", ".")
    medida = float(medida)
    medida = int(medida)

    if medida == 320:
        excel_path = "programas/202501_Spectra_List_320nm.xls"
    elif medida == 300:
        excel_path = "programas/202501_Spectra_List_300nm.xls"
    else:
        excel_path = "programas/202501_Spectra_List_280nm.xls"
    print("Plantilla elegida: ", excel_path)

    return excel_path

def copiar_datos_excel(Muestra):
    """
    Abre la plantilla de Excel y copia los datos de los archivos en las celdas correspondientes
    en la hoja "refl5".
    """
    #excel_path = os.path.abspath("programas/plantillas_macros.xls")
    excel_path = elegir_plantilla(Muestra)

    output_folder = os.path.dirname(os.path.abspath(Muestra.path_zero))
    output_folder = os.path.dirname(output_folder)  # Subir un nivel
    excel_path_output = os.path.join(output_folder, f"{Muestra.nombre}.xls")

    print(f"📂 Intentando abrir: {excel_path}")

    n_muestras = len(Muestra.lista_espect_muestras)

    try:
        wb = xw.Book(excel_path)
        sheet = wb.sheets['refl5']  # Trabajar en la hoja correcta

        # Diccionario con las celdas de inicio
        measurement_cells = {1: "Q536", 2: "R536", 3: "S536"}
        zero_line_cell = "M536"
        baseline_cell = "N536"

        def insertar_datos_en_columna(file_path, start_cell):
            """ Lee un archivo .asc y copia los valores en una columna en Excel. """
            data = leer_asc_para_exportar_excel(file_path)
            if not data:
                print(f"⚠️ No se encontraron datos en {file_path}")
                return

            start_row = int(start_cell[1:])  # Extraer el número de fila
            col_letter = start_cell[0]  # Extraer la letra de la columna

            # Escribir datos en la hoja de Excel
            for i, value in enumerate(data):
                sheet[f"{col_letter}{start_row + i}"].value = value

            print(f"✅ Copiados datos de {file_path} en columna {col_letter}{start_row}")

        # Copiar datos de los archivos de medición
        for i, file in enumerate(Muestra.lista_espect_muestras, start=1):
            if i > 3:
                print(f"⚠️ Más de 3 mediciones. Ignorando {file}")
                break
            insertar_datos_en_columna(file, measurement_cells[i])

        # Copiar datos de Zero Line y Base Line
        insertar_datos_en_columna(Muestra.path_zero, zero_line_cell)
        insertar_datos_en_columna(Muestra.path_base, baseline_cell)

        #Eliminar contenido de celdas en función de número de medidas por muestra
        if n_muestras == 2:
            sheet["C61"].value = ""
        elif n_muestras == 1:
            sheet["C61"].value = ""
            sheet["C60"].value = ""

        wb.save(excel_path_output)
        wb.close()
        print(f"📂 Archivo guardado en: {excel_path_output}")

    except Exception as e:
        print(f"❌ Error copiando datos en Excel: {e}")
