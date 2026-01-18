
import pandas as pd
import numpy as np
import re
import os



class Muestra:
    def __init__(self, nombre_muestra, archivos_ir, file_path_zero_base_uv, file_paths_muestras_uv, referencias_ir,
                 referencias_uv, datos_basicos, excel_path_output):
        self.nombre = nombre_muestra
        self.archivo_tfir = archivos_ir
        self.col_ir_ref = referencias_ir
        self.col_uv_ref = referencias_uv
        self.tipo_medida = datos_basicos["medida"]
        self.tipo = datos_basicos["aparatos"]
        self.test = datos_basicos["test"]
        self.fabricante = datos_basicos["fabricante"]
        self.proyecto = datos_basicos["proyecto"]
        self.hours = datos_basicos["hours"]
        self.meses = datos_basicos["months"]
        self.temperatura = datos_basicos["temperatura"]
        self.fechamedida = datos_basicos["fecha_medida"]["dd/mm/yyyy"]
        self.id_medida = datos_basicos["fecha_medida"]["yyyyMMdd"]
        self.path_output = excel_path_output
        if file_paths_muestras_uv:
            try:
                self.archivo_uv = {
                    "path_muestras": file_paths_muestras_uv, "zero": file_path_zero_base_uv["ZeroLine"], "base": file_path_zero_base_uv["BaseLine"],
                    "ventana": file_path_zero_base_uv["ventana"][0]}
            except:
                self.archivo_uv = {
                    "path_muestras": file_paths_muestras_uv, "zero": file_path_zero_base_uv["ZeroLine"],
                    "base": file_path_zero_base_uv["BaseLine"], "ventana": file_path_zero_base_uv["ventana"]}
        else:
            self.archivo_uv = {
                "path_muestras": "", "zero": "", "base": "", "ventana": ""}

    def leer_archivo_referencias(self):
        directorio_script = os.path.dirname(os.path.realpath(__file__))
        print(directorio_script)
        archivo_referencias = os.path.normpath(
            os.path.join(directorio_script, "../user_templates/references.xlsx"))
        print(archivo_referencias)
        return archivo_referencias

    def leer_datos_referencia(self, referencias):
        """
        Lee datos de referencia de un archivo Excel y los combina en un DataFrame.

        Args:
            referencias (dict): Diccionario con las referencias a leer, donde las claves son
                                los nombres de las columnas y los valores son los nombres
                                de las columnas en el archivo Excel.

        Returns:
            pandas.DataFrame: DataFrame combinado con los datos de referencia.
        """

        df_referencia_total = pd.DataFrame()

        for columna_nombre, columna_excel in referencias.items():
            if columna_excel:  # Verifica si la columna_excel no es None o vacío
                if columna_nombre in ['r_ftir']:
                    hoja = "absorbedores_refl"
                    col_nm = "wvl [µm]"
                elif columna_nombre == "r_trans_uv" in referencias:
                    hoja = "T_ventana_uv"
                    col_nm = "wvl [nm]"
                elif columna_nombre == "r_trans_ir" in referencias:
                    hoja = "T_ventana_ir"
                    col_nm = "wvl [nm]"
                elif columna_nombre == "r_uv":
                    hoja = "absorbedores_abs"
                    col_nm = "wvl [nm]"
                else:
                    continue  # Si no se encuentra ninguna de las anteriores, continuar.

                try:
                    archivo_referencias = self.leer_archivo_referencias()
                    data_ref = pd.read_excel(archivo_referencias, sheet_name=hoja, usecols=[col_nm, columna_excel])
                    data_ref.rename(columns={col_nm: "nm", columna_excel: columna_nombre}, inplace=True)

                    if col_nm == "wvl [µm]":
                        data_ref["nm"] = data_ref["nm"] * 1000
                        data_ref['nm'] = data_ref['nm'].round().astype(int)
                    if df_referencia_total.empty:
                        data_ref['nm'] = data_ref['nm'].round().astype(int)
                        df_referencia_total = data_ref
                    else:
                        df_referencia_total['nm'] = df_referencia_total['nm'].round().astype(int)
                        data_ref['nm'] = data_ref['nm'].round().astype(int)
                        df_referencia_total = pd.merge(df_referencia_total, data_ref, on="nm",
                                                       how='outer')  # how='outer' para que no se pierdan datos

                except FileNotFoundError:
                    print(f"Error: El archivo 'references.xlsx' no se encontró.")
                    return None
                except KeyError:
                    print(
                        f"Error: La hoja '{hoja}' o la columna '{columna_excel}' no se encontraron en el archivo Excel.")
                    return None
                except Exception as e:
                    print(f"Error inesperado al leer la hoja '{hoja}': {e}")
                    return None
        return df_referencia_total

    def procesar_datos_tfir(self):
        """
        Procesa datos de FTIR a partir del diccionario de configuración self.archivo_tfir.

        Args:
            self: Instancia de la clase que contiene el diccionario 'archivo_tfir'.

        Returns:
            pandas.DataFrame: DataFrame combinado con los datos de muestras y referencias.
        """
        archivos_ir = self.archivo_tfir
        print("archivos_ir")
        print(archivos_ir)
        # 1. Leer y procesar muestras
        df_muestras = pd.DataFrame()
        contador_muestras = 0  # Inicializa el contador

        if 'path_muestras' in archivos_ir and archivos_ir['path_muestras']:
            for ruta_archivo, lista_hojas in archivos_ir['path_muestras'].items():
                for hoja in lista_hojas:
                    try:
                        df = pd.read_excel(ruta_archivo, sheet_name=hoja, header=4)
                        df = df[['nm', '%R']]

                        # Incrementa el contador y úsalo para nombrar la columna
                        contador_muestras += 1
                        df = df.rename(columns={'%R': f'I{contador_muestras}'})

                        if df_muestras.empty:
                            df_muestras = df.copy()
                        else:
                            df_muestras = pd.merge(df_muestras, df, on='nm', how='outer')
                    except Exception as e:
                        print(f"Error al leer la hoja '{hoja}' del archivo '{ruta_archivo}': {e}")
        else:
            print("No se encontraron datos de muestras para procesar.")

        # 2. Leer y procesar referencias
        df_referencias = pd.DataFrame()
        #referencia_keys = ['zero', 'baseoro', 'basenegro', 'ventana', 'ventanaoro',
                           #'ventananegro']  # lista de claves de ref
        referencia_keys = ['zero', 'base', 'ventana']  # lista de claves de ref
        for ref_key in referencia_keys:
            if ref_key in archivos_ir and archivos_ir[ref_key]:  # verifica si la clave existe y no esta vacia
                for ruta_archivo, hoja in archivos_ir[ref_key].items():
                    try:
                        df = pd.read_excel(ruta_archivo, sheet_name=hoja, header=4)
                        df = df[['nm', '%R']]
                        df = df.rename(columns={'%R': ref_key})  # Usa el nombre de la clave como nombre de columna
                        if df_referencias.empty:
                            df_referencias = df.copy()  # usa copy
                        else:
                            df_referencias = pd.merge(df_referencias, df, on='nm', how='outer')  # outer join
                    except Exception as e:
                        print(f"Error al leer la hoja '{hoja}' del archivo '{ruta_archivo}': {e}")
            else:
                print(f"No se encontraron datos de referencia para '{ref_key}'.")

        # 3. Combinar muestras y referencias
        if not df_muestras.empty and not df_referencias.empty:
            df_final = pd.merge(df_muestras, df_referencias, on='nm', how='outer')  # Usa outer join
        elif not df_muestras.empty:
            df_final = df_muestras
        elif not df_referencias.empty:
            df_final = df_referencias
        else:
            df_final = pd.DataFrame()  # Devuelve un DataFrame vacío si no hay datos

        # 4. Combinar con datos de referencia total
        datos_referencia_total = self.leer_datos_referencia(
            self.col_ir_ref)  # Asumiendo que este metodo existe en la clase

        if datos_referencia_total is not None and not df_final.empty:
            datos_referencia_total['nm'] = datos_referencia_total['nm'].round().astype(int)
            df_final['nm'] = df_final['nm'].astype(int)
            df_final = pd.merge(df_final, datos_referencia_total, on='nm', how='left')
        elif datos_referencia_total is not None:
            df_final = datos_referencia_total

        print("df ftir")
        print(df_final)
        return df_final

    def leer_asc(selfself, path_asc):
        with open(path_asc, 'r') as file:
            lines = file.readlines()
        return lines
    def leer_datos_asc_columnas(self, path_asc, col_name):
        with open(path_asc, 'r') as file:
            lines = file.readlines()

        data_start = None
        for i, line in enumerate(lines):
            if '#DATA' in line:
                data_start = i + 1
                break

        if data_start is None:
            print("La sección #DATA no se encontró en el archivo.")
            return None

        data_lines = lines[data_start:]
        data = [line.strip().replace(',', '.').split('\t') for line in data_lines]
        df = pd.DataFrame(data, columns=['nm', col_name]).astype(float)

        return df

    def leer_datos_UV(self, ventana):
        print(f"Leyendo datos UV de la muestra {self.nombre}")

        data_uv = pd.DataFrame()
        j = 1

        # Leer muestras
        for path_muestra in self.archivo_uv["path_muestras"]:
            df = self.leer_datos_asc_columnas(path_muestra, f"I{j}")
            if data_uv.empty:
                data_uv = df
            else:
                data_uv = pd.merge(data_uv, df, on="nm", how="inner")
            j += 1

        # Leer zero, base
        zero_df = self.leer_datos_asc_columnas(self.archivo_uv["zero"], "zero")
        base_df = self.leer_datos_asc_columnas(self.archivo_uv["base"], "base")

        # Combinar datos
        data_uv = pd.merge(data_uv, zero_df, on="nm", how="inner")
        data_uv = pd.merge(data_uv, base_df, on="nm", how="inner")

        # Leer ventana y ventanabase si no son None
        if ventana == True:
            ventana_df = self.leer_datos_asc_columnas(self.archivo_uv["ventana"], "ventana")
            data_uv = pd.merge(data_uv, ventana_df, on="nm", how="inner")
            #ventanabase_df = self.leer_datos_asc_columnas(self.archivo_uv["ventanabase"], "ventanabase")
            #data_uv = pd.merge(data_uv, ventanabase_df, on="nm", how="inner")

        # Leer datos de referencia
        data_ref = self.leer_datos_referencia(self.col_uv_ref)
        archivo_referencias = self.leer_archivo_referencias()
        data_espectro = pd.read_excel(archivo_referencias, sheet_name="absorbedores_espectro", usecols=["wvl [nm]", "ASTM G173 direct"])
        data_espectro.rename(columns={"ASTM G173 direct": "ASTM", "wvl [nm]": "nm"}, inplace= True)
        # Combinar datos de referencia
        data_uv = pd.merge(data_uv, data_ref, on="nm", how="inner")
        data_uv = pd.merge(data_uv, data_espectro, on="nm", how="inner")

        data_uv.sort_values(by="nm", inplace=True)

        return data_uv

    def medidas_UV(self, df, ventana):
        print("Calculando medidas UV")
        # Identificar las columnas de Intensidad dinámicamente usando una expresión regular
        intensity_cols = [col for col in df.columns if re.match(r'^I\d+$', col)]

        # Calcular la media solo si existen columnas de intensidad
        if intensity_cols:
            df['Iw'] = df[intensity_cols].mean(axis=1)

            refl_list_std = []
            for intensity in intensity_cols:
                reflectance = (((df[intensity] -df["zero"]) / (df["base"] - df["zero"])) * df["r_uv"])*df["ASTM"].sum()
                refl_list_std.append(reflectance)
            SWR_std = np.std(refl_list_std)

        else:
            df['Iw'] = np.nan  # O algún valor por defecto si no hay columnas I

        if ventana == False:
            df["refl"] = ((df["Iw"] - df["zero"]) / (df["base"] - df["zero"])) * df["r_uv"]
            df["abs"] = 1 - df["refl"]
        else:
            df["refl_ventana"] = ((df["ventana"]-df["zero"])/(df["base"] - df["zero"]))*df["r_uv"]
            df["refl_muestras"] =(df["Iw"]-df["zero"])/(df["base"]-df["zero"])*df["r_uv"]
            df["refl"] = ((df["refl_muestras"]-df["refl_ventana"])/
                          (((df["r_trans_uv"]/100)**2)+df["refl_ventana"]*(df["refl_muestras"]-df["refl_ventana"])))
            df["abs"] = 1- df["refl"]
        SWR = (df['refl'] * df['ASTM']).sum()
        SWA = (df["abs"] * df["ASTM"]).sum()

        df_output = df[["nm", "refl", "abs"]]

        return df_output, SWR, SWA, SWR_std

    def medidas_ir(self, df, ventana):
        print("Calculando medidas IR")
        # Identificar las columnas de Intensidad dinámicamente usando una expresión regular
        intensity_cols = [col for col in df.columns if re.match(r'^I\d+$', col)]

        # Calcular la media solo si existen columnas de intensidad
        if intensity_cols:
            df['Iw'] = df[intensity_cols].mean(axis=1)
        else:
            df['Iw'] = np.nan  # O algún valor por defecto si no hay columnas I

        if ventana == False:
            df["refl"] = (df["Iw"] - df["zero"]) / (df["base"] - df["zero"]) * df["r_ftir"]
            df["abs"] = 1 - df["refl"]
        else:
            df["refl_ventana"] = ((df["ventana"] - df["zero"]) / (df["base"] - df["zero"])) * df["r_ftir"]
            df["refl_muestras"] = (df["Iw"] - df["zero"]) / (df["base"] - df["zero"]) * df["r_ftir"]
            df["refl"] = ((df["refl_muestras"] - df["refl_ventana"]) /
                          (((df["r_trans_ir"] / 100) ** 2) + df["refl_ventana"] * (df["refl_muestras"] - df["refl_ventana"])))
            df["abs"] = 1 - df["refl"]

        df_output = df[["nm", "refl", "abs"]]

        return df_output

    def combinar_uv_ir(self, data_ir, data_uv):

        if not data_ir.empty and not data_uv.empty:
            data_ir = data_ir.loc[data_ir["nm"] > 2500]
            data_ir = data_ir.loc[data_ir["nm"] <= 16000]
            df_concatenado = pd.concat([data_ir, data_uv], ignore_index=True).sort_values(by='nm')
        elif not data_ir.empty:
            data_ir = data_ir.loc[data_ir["nm"] > 2500]
            data_ir = data_ir.loc[data_ir["nm"] <= 16000]
            df_concatenado = data_ir.sort_values(by='nm')
            df_concatenado["µm"] = df_concatenado["nm"]/1000
        else:
            df_concatenado = data_uv.sort_values(by='nm')
            df_concatenado["µm"] = df_concatenado["nm"]/1000

        return df_concatenado


    def emitancia(self, df):

        df["µm"] = df["nm"]/1000
        temp = float(self.temperatura)
        ### M_bb #####
        # Constantes físicas
        h = 6.63e-34  # Constante de Planck (J·s)
        c = 2.997e8  # Velocidad de la luz (m/s)
        k_B = 1.38e-23  # Constante de Boltzmann (J/K)

        # Temperatura (en Kelvin)
        T = temp  # Define temp antes de ejecutar el código

        # Cargar datos de un DataFrame (supongamos que ya lo tienes cargado en df)
        df["lambda_m"] = df["µm"] * 1e-6  # Convertir micrómetros a metros

        # Calcular M_bb para cada valor de lambda
        df["M_bb"] = (2 * np.pi * h * c ** 2) / (df["lambda_m"] ** 5) * \
                     (1 / (np.exp((h * c) / (df["lambda_m"] * k_B * T)) - 1))

        df["integrando"] = df["M_bb"] * df["abs"] * df["lambda_m"].diff()

        df["denominador"] = df["M_bb"] * df["lambda_m"].diff()

        emitancia = (df["integrando"].sum()) / (df["denominador"].sum())

        return emitancia, df
