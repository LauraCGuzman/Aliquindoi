import xlwings as xw
import os
import string
from datetime import datetime
from openpyxl.utils import get_column_letter


def leer_asc_para_exportar_excel_antiguo(path_asc):
    with open(path_asc, 'r') as file:
        lines = file.readlines()

    data_start = False
    data = []
    # hacer esta parte más robusta
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


def leer_asc_para_exportar_excel(path_asc):
    with open(path_asc, 'r') as file:
        lines = file.readlines()

    data_start = False
    data = []
    # hacer esta parte más robusta
    for i, line in enumerate(lines):
        if "#DATA" in line:
            data_start = True
        if  data_start:
            split_line = line.split()
            if len(split_line) > 1:  # Asegurar que hay al menos dos elementos
                cleaned_value = split_line[1].replace(",", ".")  # Reemplazar , por .
                data.append(cleaned_value)  # Guardar la segunda columna ya corregida

    return data

def elegir_plantilla(Muestra):

    def elegir_plantilla_segun_nm(Muestra):
        if Muestra.tipo_medida == "Transmitancia CSP":
            path = Muestra.archivo_uv["path_muestras"][0] # no hay base o zero, hay que fijarse en el primer archivo muestra
        else:
            path = Muestra.archivo_uv["zero"] #para el resto sí hay zero.
        with open(path, 'r') as file:
            lines = file.readlines()

        linea = lines[-1]
        split_line = linea.split()
        medida_nm = split_line[0].replace(",", ".")
        medida_nm = float(medida_nm)
        medida_nm = int(medida_nm)

        return medida_nm

    directorio_script = os.path.dirname(os.path.realpath(__file__))

    if Muestra.tipo_medida == "Reflectancia":
        medida_nm = elegir_plantilla_segun_nm(Muestra)
        if medida_nm == 320:
            excel_path = os.path.normpath(
                os.path.join(directorio_script, "../para_el_usuario/202501_refl_320.xls"))
        elif medida_nm == 300:
            excel_path = os.path.normpath(
                os.path.join(directorio_script, "../para_el_usuario/202501_refl_300.xls"))
        else:
            excel_path = os.path.normpath(
                os.path.join(directorio_script, "../para_el_usuario/202501_refl_280.xls"))
        print("Plantilla elegida: ", excel_path)
    elif Muestra.tipo_medida == "Transmitancia CSP":
        medida_nm = elegir_plantilla_segun_nm(Muestra)
        if medida_nm == 320:
            excel_path = os.path.normpath(
                os.path.join(directorio_script, "../para_el_usuario/202501_transcsp_320.xls"))
        elif medida_nm == 300:
            excel_path = os.path.normpath(
                os.path.join(directorio_script, "../para_el_usuario/202501_transcsp_300.xls"))
        else:
            excel_path = os.path.normpath(
                os.path.join(directorio_script, "../para_el_usuario/202501_transcsp_280.xls"))
        print("Plantilla elegida: ", excel_path)
    elif Muestra.tipo_medida == "Transmitancia PV":
        excel_path = os.path.normpath(
            os.path.join(directorio_script, "../para_el_usuario/2025_transpv_320nm.xls"))
    elif Muestra.tipo_medida == "Absortancia":
        medida_nm = elegir_plantilla_segun_nm(Muestra)
        if medida_nm == 320:
            excel_path = os.path.normpath(
                os.path.join(directorio_script, "../para_el_usuario/abs_plantilla_320.xlsx"))
        elif medida_nm == 300:
            excel_path = os.path.normpath(
                os.path.join(directorio_script, "../para_el_usuario/abs_plantilla_300.xlsx"))
        else:
            excel_path = os.path.normpath(
                os.path.join(directorio_script, "../para_el_usuario/abs_plantilla_280.xlsx"))
        print("Plantilla elegida: ", excel_path)
    return excel_path

def copiar_datos_excel(Muestra, wb_destino):
    """
    Abre la plantilla, modifica la hoja "refl5" y la guarda en el libro de trabajo de destino.
    """

    plantilla_path = elegir_plantilla(Muestra)

    print(f" Intentando abrir plantilla: {plantilla_path}")
    print(f" Intentando abrir libro de destino: {Muestra.path_output}")
    print("Hecho")

    n_muestras = len(Muestra.archivo_uv["path_muestras"])

    try:
        wb_plantilla = xw.Book(plantilla_path)  # Plantilla

        app = wb_destino.app  # Obtener la aplicación de Excel
        app.display_alerts = False  # Desactivar alertas


        # Diccionario con las celdas de inicio
        if Muestra.tipo_medida == "Reflectancia":
            measurement_cells = {1: "Q628", 2: "R628", 3: "S628", 4:"T628", 5: "U628", 6: "V628", 7:"W628", 8:"X628", 9:"Y628"}
            zero_line_cell = "M628"
            baseline_cell = "N628"
            sheet_plantilla = wb_plantilla.sheets['refl5']  # Hoja de la plantilla

        elif Muestra.tipo_medida == "Transmitancia CSP":
            measurement_cells = {1: "T628", 2: "U628", 3: "V628", 4: "W628", 5: "X628", 6: "Y628", 7: "Z628", 8: "AA628",
                                 9: "AB628"}
            zero_line_cell = "M628"
            baseline_cell = "N628"
            sheet_plantilla = wb_plantilla.sheets['trans']  # Hoja de la plantilla

        elif Muestra.tipo_medida == "Transmitancia PV":
            measurement_cells = {1: "Q628", 2: "R628", 3: "S628", 4: "T628", 5: "U628", 6: "V628", 7: "W628", 8: "X628",
                                 9: "Y628"}
            zero_line_cell = "M628"
            baseline_cell = "N628"
            sheet_plantilla = wb_plantilla.sheets['trans5_PV_5nm']  # Hoja de la plantilla

        print("Determinadas celdas para copiar datos")

        def col_letter_to_index(col_letter):
            """ Convierte una letra de columna (A, B, ..., Z, AA, AB, ...) a un índice numérico """
            col_letter = col_letter.upper()
            col_index = 0
            for char in col_letter:
                col_index = col_index * 26 + (string.ascii_uppercase.index(char) + 1)
            return col_index

        def insertar_datos_en_columna(file_path, start_cell, sheet):
            """ Lee un archivo .asc y copia los valores en una columna en Excel. """
            data = leer_asc_para_exportar_excel(file_path)
            if not data:
                print(f"⚠️ No se encontraron datos en {file_path}")
                return

            start_row = int(start_cell[1:])  # Extraer el número de fila
            col_letter = start_cell[0]  # Extraer la letra de la columna
            col_index = col_letter_to_index(col_letter)  # Convertir a índice numérico

            # Convertir los datos en una lista de listas para escritura en bloque
            data_array = [[value] for value in data]

            # Escribir los datos de una sola vez en Excel
            sheet.range((start_row, col_index)).value = data_array

            print(f"✅ Copiados datos de {file_path} en columna {col_letter}{start_row}")

        # Copiar datos de los archivos de medición
        for i, file in enumerate(Muestra.archivo_uv["path_muestras"], start=1):
            if i > 9:
                print(f"⚠️ Más de 9 mediciones. Ignorando {file}")
                break
            insertar_datos_en_columna(file, measurement_cells[i], sheet_plantilla)

        # Copiar datos de Zero Line y Base Line
        if Muestra.tipo_medida == "Reflectancia":
            insertar_datos_en_columna(Muestra.archivo_uv["zero"], zero_line_cell, sheet_plantilla)
            insertar_datos_en_columna(Muestra.archivo_uv["base"], baseline_cell, sheet_plantilla)

        #Eliminar contenido de celdas en función de número de medidas por muestra
        j = 0
        for i in range(n_muestras, 9):
            # La celda es C(67 - (índice de desfase))
            celda = f"C{67 - j}"

            sheet_plantilla[celda].value = ""
            print("eliminando celda ", celda)
            j = j+1

        #introducir datos
        sheet_plantilla["C3"].value = Muestra.nombre + " - " + Muestra.fabricante
        sheet_plantilla["C5"].value = Muestra.nombre
        sheet_plantilla["C6"].value = Muestra.nombre
        sheet_plantilla["C7"].value = Muestra.fabricante
        sheet_plantilla["C10"].value = Muestra.proyecto
        sheet_plantilla["C11"].value = datetime.strptime(Muestra.fechamedida, "%d/%m/%Y")
        sheet_plantilla["C12"].value = Muestra.id_medida
        sheet_plantilla["C15"].value = Muestra.test
        if Muestra.tipo_medida == "Reflectancia":
            sheet_plantilla["C21"].value = Muestra.hours
            sheet_plantilla["C22"].value = Muestra.meses
            sheet_plantilla["C23"].value = str(Muestra.col_uv_ref["r_uv"])
        else:
            sheet_plantilla["C21"].value = Muestra.meses


        # Copiar la hoja sin renombrarla
        new_sheet = sheet_plantilla.copy(after=wb_destino.sheets[wb_destino.sheets.count - 1])

        # Renombrar la hoja copiada
        sheet_name  = Muestra.nombre + "_" + Muestra.id_medida
        new_sheet.name = sheet_name[:31]

        wb_plantilla.close()

        print(f" Archivo guardado en: {Muestra.path_output}")

    except Exception as e:
        print(f"❌ Error copiando datos en Excel: {e}")

def copiar_datos_excel_absorbedores(Muestra, df, wb_destino, SWR_uv, SWA_uv, SWR_std, emitancia, temperatura,
                                    dataframe_ir, dataframe_uv, df_abs):

    plantilla_path = elegir_plantilla(Muestra)

    print(f" Intentando abrir plantilla: {plantilla_path}")
    print(f" Intentando abrir libro de destino: {Muestra.path_output}")

    #try:
    wb_plantilla = xw.Book(plantilla_path)  # Plantilla

    app = wb_destino.app  # Obtener la aplicación de Excel
    app.display_alerts = False  # Desactivar alertas

    sheet_plantilla = wb_plantilla.sheets['Results']  # Hoja de la plantilla

    # introducir datos
    sheet_plantilla["I1"].value = Muestra.proyecto
    sheet_plantilla["I2"].value = Muestra.nombre + " - " + Muestra.fabricante
    sheet_plantilla["K5"].value = Muestra.nombre
    sheet_plantilla["M1"].value = Muestra.fabricante
    sheet_plantilla["K1"].value = datetime.strptime(Muestra.fechamedida, "%d/%m/%Y")
    sheet_plantilla["K2"].value = Muestra.id_medida
    sheet_plantilla["I3"].value = Muestra.test
    sheet_plantilla["I4"].value = Muestra.hours
    sheet_plantilla["I5"].value = Muestra.meses
    sheet_plantilla["K3"].value = str(Muestra.col_uv_ref["r_uv"])
    sheet_plantilla["I7"].value = SWR_uv
    sheet_plantilla["I8"].value = SWA_uv
    sheet_plantilla["I9"].value = temperatura
    sheet_plantilla["I11"].value = SWR_std

    if "FTIR" in Muestra.tipo and Muestra.archivo_tfir:
        sheet_plantilla["K4"].value = str(Muestra.col_ir_ref["r_ftir"])
    else:
        sheet_plantilla["I10"].value = ""
        sheet_plantilla.range("Q1:S3824").value = None

    sheet_plantilla.range("A5").options(index=False, header=False).value = df["µm"].values.reshape(-1, 1)
    sheet_plantilla.range("B5").options(index=False, header=False).value = df["refl"].values.reshape(-1, 1)
    sheet_plantilla.range("C5").options(index=False, header=False).value = df["abs"].values.reshape(-1, 1)

    if "FTIR" in Muestra.tipo and Muestra.archivo_tfir:
        sheet_plantilla.range("X2").options(index=False, header=True).value = [dataframe_ir.columns.tolist()] + dataframe_ir.values.tolist()
    sheet_plantilla.range("AQ2").options(index=False, header=True).value = [dataframe_uv.columns.tolist()] + dataframe_uv.values.tolist()

    # Copiar la hoja sin renombrarla
    new_sheet = sheet_plantilla.copy(after=wb_destino.sheets[wb_destino.sheets.count - 1])

    # Renombrar la hoja copiada
    sheet_name = Muestra.nombre + "_" + Muestra.id_medida
    new_sheet.name = sheet_name[:31]

    wb_plantilla.close()

    print(f" Archivo guardado en: {Muestra.path_output}")

