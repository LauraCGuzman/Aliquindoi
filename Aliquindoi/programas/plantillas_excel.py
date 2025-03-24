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
    path = Muestra.archivo_uv["zero"]
    with open(path, 'r') as file:
        lines = file.readlines()

    linea = lines[-1]
    split_line = linea.split()
    medida_nm = split_line[0].replace(",", ".")
    medida_nm = float(medida_nm)
    medida_nm = int(medida_nm)

    if Muestra.tipo_medida == "Reflectancia":
        if medida_nm == 320:
            excel_path = "programas/202501_refl_320.xls"
        elif medida_nm == 300:
            excel_path = "programas/202501_refl_300.xls"
        else:
            excel_path = "programas/202501_refl_280.xls"
        print("Plantilla elegida: ", excel_path)
    elif Muestra.tipo_medida == "Transmitancia CSP":
        if medida_nm == 320:
            excel_path = "programas/202501_transcsp_320.xls"
        elif medida_nm == 300:
            excel_path = "programas/202501_transcsp_300.xls"
        else:
            excel_path = "programas/202501_transcsp_280.xls"
        print("Plantilla elegida: ", excel_path)
    elif Muestra.tipo_medida == "Transmitancia PV":
        excel_path = "programas/2025_trans_320nm_PV.xls"

    return excel_path

def copiar_datos_excel(Muestra, wb_destino):
    """
    Abre la plantilla, modifica la hoja "refl5" y la guarda en el libro de trabajo de destino.
    """

    plantilla_path = elegir_plantilla(Muestra)

    print(f" Intentando abrir plantilla: {plantilla_path}")
    print(f" Intentando abrir libro de destino: {Muestra.path_output}")

    n_muestras = len(Muestra.lista_espect_muestras)

    try:
        wb_plantilla = xw.Book(plantilla_path)  # Plantilla

        app = wb_destino.app  # Obtener la aplicación de Excel
        app.display_alerts = False  # Desactivar alertas

        sheet_plantilla = wb_plantilla.sheets['refl5']  # Hoja de la plantilla

        # Diccionario con las celdas de inicio
        measurement_cells = {1: "Q536", 2: "R536", 3: "S536"}
        zero_line_cell = "M536"
        baseline_cell = "N536"

        def insertar_datos_en_columna(file_path, start_cell, sheet):
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
            insertar_datos_en_columna(file, measurement_cells[i], sheet_plantilla)

        # Copiar datos de Zero Line y Base Line
        insertar_datos_en_columna(Muestra.path_zero, zero_line_cell, sheet_plantilla)
        insertar_datos_en_columna(Muestra.path_base, baseline_cell, sheet_plantilla)

        #Eliminar contenido de celdas en función de número de medidas por muestra
        if n_muestras == 2:
            sheet_plantilla["C61"].value = ""
        elif n_muestras == 1:
            sheet_plantilla["C61"].value = ""
            sheet_plantilla["C60"].value = ""

        #introducir datos
        sheet_plantilla["C3"].value = Muestra.nombre + " - " + Muestra.fabricante
        sheet_plantilla["C5"].value = Muestra.nombre
        sheet_plantilla["C6"].value = Muestra.nombre
        sheet_plantilla["C7"].value = Muestra.fabricante
        sheet_plantilla["C11"].value = Muestra.fechamedida
        sheet_plantilla["C12"].value = Muestra.id_medida
        sheet_plantilla["C15"].value = Muestra.test
        sheet_plantilla["C21"].value = Muestra.hours
        sheet_plantilla["C22"].value = Muestra.meses
        sheet_plantilla["C23"].value = str(Muestra.col_uv_ref["r_uv"])

        # Copiar la hoja sin renombrarla
        new_sheet = sheet_plantilla.copy(after=wb_destino.sheets[wb_destino.sheets.count - 1])

        # Renombrar la hoja copiada
        new_sheet.name = Muestra.nombre + "_" + Muestra.id_medida

        wb_plantilla.close()

        print(f" Archivo guardado en: {Muestra.path_output}")

    except Exception as e:
        print(f"❌ Error copiando datos en Excel: {e}")

def copiar_datos_excel_absorbedores(Muestra, df, wb_destino, SWR_uv, SWA_uv, emitancia, temperatura,
                                    dataframe_ir, dataframe_uv):

    plantilla_path = "programas/abs_plantilla.xlsx"

    print(f" Intentando abrir plantilla: {plantilla_path}")
    print(f" Intentando abrir libro de destino: {Muestra.path_output}")

    try:
        wb_plantilla = xw.Book(plantilla_path)  # Plantilla

        app = wb_destino.app  # Obtener la aplicación de Excel
        app.display_alerts = False  # Desactivar alertas

        sheet_plantilla = wb_plantilla.sheets['Results']  # Hoja de la plantilla

        # introducir datos
        sheet_plantilla["I1"].value = Muestra.proyecto
        sheet_plantilla["I2"].value = Muestra.nombre + " - " + Muestra.fabricante
        sheet_plantilla["K5"].value = Muestra.nombre
        sheet_plantilla["M1"].value = Muestra.fabricante
        sheet_plantilla["K1"].value = Muestra.fechamedida
        sheet_plantilla["K2"].value = Muestra.id_medida
        sheet_plantilla["I3"].value = Muestra.test
        sheet_plantilla["I4"].value = Muestra.hours
        sheet_plantilla["I5"].value = Muestra.meses
        sheet_plantilla["K3"].value = str(Muestra.col_uv_ref["r_uv"])
        sheet_plantilla["K4"].value = str(Muestra.col_ir_ref["r_oro"])+", "+str(Muestra.col_ir_ref["r_negro"])

        sheet_plantilla["I7"].value = SWR_uv
        sheet_plantilla["I8"].value = SWA_uv
        sheet_plantilla["I9"].value = temperatura
        #sheet_plantilla["I10"].value = emitancia

        sheet_plantilla.range("A5").options(index=False, header=False).value = df["µm"].values.reshape(-1, 1)
        sheet_plantilla.range("B5").options(index=False, header=False).value = df["refl"].values.reshape(-1, 1)
        sheet_plantilla.range("C5").options(index=False, header=False).value = df["abs"].values.reshape(-1, 1)

        sheet_plantilla.range("X2").options(index=False, header=True).value = [dataframe_ir.columns.tolist()] + dataframe_ir.values.tolist()
        sheet_plantilla.range("AQ2").options(index=False, header=True).value = [dataframe_uv.columns.tolist()] + dataframe_uv.values.tolist()

        # Copiar la hoja sin renombrarla
        new_sheet = sheet_plantilla.copy(after=wb_destino.sheets[wb_destino.sheets.count - 1])

        # Renombrar la hoja copiada
        new_sheet.name = Muestra.nombre + "_" + Muestra.id_medida

        wb_plantilla.close()

        print(f" Archivo guardado en: {Muestra.path_output}")

    except Exception as e:
        print(f"❌ Error copiando datos en Excel: {e}")