import xlwings as xw
import os
import string
import ast
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string

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

def elegir_plantilla_config(Muestra, config):

    directorio_script = os.path.dirname(os.path.realpath(__file__))

    # Select key based on measurement type
    if Muestra.tipo_medida == "Reflectancia":
        key = "reflectance"
    elif Muestra.tipo_medida == "Transmitancia CSP":
        key = "transmittance_csp"
    elif Muestra.tipo_medida == "Transmitancia PV":
        key = "transmittance_pv"
    elif Muestra.tipo_medida == "Absortancia":
        key = "absorbance"
    else:
        print(f"⚠️ Tipo de medida no reconocido: {Muestra.tipo_medida}. Usando Reflectancia por defecto.")
        key = "reflectance"

    plantilla_rel_path = config.path_plantillas.get(key)
    if not plantilla_rel_path:
        print(f"❌ No se encontró configuración para '{key}' en config.ini")
        return None, {}

    excel_path = os.path.normpath(
        os.path.join(directorio_script, str(plantilla_rel_path)))
    celdas = config.celdas_plantillas.get(key, {})

    print("Plantilla elegida: ", excel_path)
    return excel_path, celdas

def copiar_datos_excel(Muestra, wb_destino, config):
    """
    Abre la plantilla, modifica la hoja y la guarda en el libro de trabajo de destino.
    """

    plantilla_path, celdas = elegir_plantilla_config(Muestra, config)

    if not plantilla_path:
        print("❌ No se pudo determinar la plantilla.")
        return

    # Parse measurement_cells safely if it is a string (from configparser)
    if "measurement_cells" in celdas and isinstance(celdas["measurement_cells"], str):
        try:
             celdas["measurement_cells"] = ast.literal_eval(celdas["measurement_cells"])
        except Exception as e:
             print(f"Error parsing measurement_cells: {e}")

    print(f" Intentando abrir plantilla: {plantilla_path}")
    print(f" Intentando abrir libro de destino: {Muestra.path_output}")
    print("Hecho")

    n_muestras = len(Muestra.archivo_uv["path_muestras"])

    try:
        wb_plantilla = xw.Book(plantilla_path)  # Plantilla

        app = wb_destino.app  # Obtener la aplicación de Excel
        app.display_alerts = False  # Desactivar alertas


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
            if "measurement_cells" in celdas and i in celdas["measurement_cells"]:
                 insertar_datos_en_columna(file, celdas["measurement_cells"][i], wb_plantilla.sheets[str(celdas["hoja_plantilla"])])
            else:
                 print(f"⚠️ No hay configuración de celda para la medición {i}")

        # Copiar datos de Zero Line y Base Line
        if Muestra.tipo_medida == "Reflectancia":
            insertar_datos_en_columna(Muestra.archivo_uv["zero"], str(celdas["zero_cell"]), wb_plantilla.sheets[str(celdas["hoja_plantilla"])])
            insertar_datos_en_columna(Muestra.archivo_uv["base"], str(celdas["base_cell"]), wb_plantilla.sheets[str(celdas["hoja_plantilla"])])

        #Eliminar contenido de celdas en función de número de medidas por muestra
        # 1. Recuperas el valor "C67" de tu variable
        if "ultima_celda_medidas_borrar" in celdas:
            ref_config = celdas["ultima_celda_medidas_borrar"]

            # 2. Separas la columna ('C') y la fila (67)
            columna, fila_base = coordinate_from_string(ref_config)

            sheet_plantilla = wb_plantilla.sheets[str(celdas["hoja_plantilla"])]

            j = 0

            max_medidas = int(celdas.get("numero_max_medidas", 9))
            for i in range(n_muestras, max_medidas):
                # 3. Calculamos la nueva fila usando la base extraída del config
                nueva_fila = fila_base - j

                # 4. Reconstruimos el string de la celda (ej: "C" + "66")
                celda = f"{columna}{nueva_fila}"

                sheet_plantilla[celda].value = ""
                print("eliminando celda ", celda)
                j = j + 1

        #introducir datos
        sheet_plantilla = wb_plantilla.sheets[str(celdas["hoja_plantilla"])]
        
        sheet_plantilla[str(celdas["celda_nombre_muestra_fabricante"])].value = Muestra.nombre + " - " + Muestra.fabricante
        sheet_plantilla[str(celdas["celda_muestra_nombre"])].value = Muestra.nombre
        sheet_plantilla[str(celdas["celda_muestra_nombre_2"])].value = Muestra.nombre
        sheet_plantilla[str(celdas["celda_muestra_fabricante"])].value = Muestra.fabricante
        sheet_plantilla[str(celdas["celda_muestra_proyecto"])].value = Muestra.proyecto
        sheet_plantilla[str(celdas["celda_fecha"])].value = datetime.strptime(Muestra.fechamedida, "%d/%m/%Y")
        sheet_plantilla[str(celdas["celda_id_medida"])].value = Muestra.id_medida
        sheet_plantilla[str(celdas["celda_test"])].value = Muestra.test
        if Muestra.tipo_medida == "Reflectancia":
            sheet_plantilla[str(celdas["celda_horas"])].value = Muestra.hours
            sheet_plantilla[str(celdas["celda_meses"])].value = Muestra.meses
            sheet_plantilla[str(celdas["celda_ref_uv"])].value = str(Muestra.col_uv_ref["r_uv"])
        else:
            sheet_plantilla[str(celdas["celda_horas"])].value = Muestra.meses


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
                                    dataframe_ir, dataframe_uv, df_abs, config):

    plantilla_path, celdas = elegir_plantilla_config(Muestra, config)
    if not plantilla_path:
        print("❌ No se pudo determinar la plantilla.")
        return

    print("Celdas")
    print(celdas)
    print(f" Intentando abrir plantilla: {plantilla_path}")
    print(f" Intentando abrir libro de destino: {Muestra.path_output}")

    #try:
    wb_plantilla = xw.Book(plantilla_path)  # Plantilla

    app = wb_destino.app  # Obtener la aplicación de Excel
    app.display_alerts = False  # Desactivar alertas

    sheet_plantilla = wb_plantilla.sheets[str(celdas["hoja_plantilla"])]  # Hoja de la plantilla

    # introducir datos
    sheet_plantilla[str(celdas["celda_muestra_proyecto"])].value = Muestra.proyecto
    sheet_plantilla[str(celdas["celda_nombre_muestra_fabricante"])].value = Muestra.nombre + " - " + Muestra.fabricante
    sheet_plantilla[str(celdas["celda_muestra_nombre"])].value = Muestra.nombre
    sheet_plantilla[str(celdas["celda_muestra_fabricante"])].value = Muestra.fabricante
    sheet_plantilla[str(celdas["celda_fecha"])].value = datetime.strptime(Muestra.fechamedida, "%d/%m/%Y")
    sheet_plantilla[str(celdas["celda_id_medida"])].value = Muestra.id_medida
    sheet_plantilla[str(celdas["celda_test"])].value = Muestra.test
    sheet_plantilla[str(celdas["celda_horas"])].value = Muestra.hours
    sheet_plantilla[str(celdas["celda_meses"])].value = Muestra.meses
    try:
        sheet_plantilla[str(celdas["columna_uv_ref"])].value = str(Muestra.col_uv_ref["r_uv"])
    except:
        pass
    sheet_plantilla[str(celdas["swr_uv"])].value = SWR_uv
    sheet_plantilla[str(celdas["swa_uv"])].value = SWA_uv
    sheet_plantilla[str(celdas["temperatura"])].value = temperatura
    sheet_plantilla[str(celdas["swr_desviacion_estandar"])].value = SWR_std

    if "FTIR" in Muestra.tipo and Muestra.archivo_tfir:
        sheet_plantilla[str(celdas["columna_ir_ref"])].value = str(Muestra.col_ir_ref["r_ftir"])
    else:
        sheet_plantilla[str(celdas["emitancia"])].value = ""
        sheet_plantilla.range("Q1:S3824").value = None

    sheet_plantilla.range(str(celdas["longitud_onda"])).options(index=False, header=False).value = df["µm"].values.reshape(-1, 1)
    sheet_plantilla.range(str(celdas["reflectancia"])).options(index=False, header=False).value = df["refl"].values.reshape(-1, 1)
    sheet_plantilla.range(str(celdas["absorptancia"])).options(index=False, header=False).value = df["abs"].values.reshape(-1, 1)

    if "FTIR" in Muestra.tipo and Muestra.archivo_tfir:
        sheet_plantilla.range(str(celdas["tabla_ir"])).options(index=False, header=True).value = [dataframe_ir.columns.tolist()] + dataframe_ir.values.tolist()
    if "Espectrofotómetro" in Muestra.tipo and Muestra.archivo_uv:
        sheet_plantilla.range(str(celdas["tabla_uv"])).options(index=False, header=True).value = [dataframe_uv.columns.tolist()] + dataframe_uv.values.tolist()

    # Copiar la hoja sin renombrarla
    new_sheet = sheet_plantilla.copy(after=wb_destino.sheets[wb_destino.sheets.count - 1])

    # Renombrar la hoja copiada
    sheet_name = Muestra.nombre + "_" + Muestra.id_medida
    new_sheet.name = sheet_name[:31]

    wb_plantilla.close()

    print(f" Archivo guardado en: {Muestra.path_output}")
