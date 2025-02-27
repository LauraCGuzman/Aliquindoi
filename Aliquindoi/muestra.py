import pandas as pd
import numpy as np
import re
class Muestra:
    def __init__(self, nombre_muestra, archivos_ir, file_path_zero_base_uv, file_paths_muestras_uv, referencias_ir,
                 referencias_uv, datos_basicos, excel_path_output):
        self.nombre = nombre_muestra
        self.archivo_tfir = archivos_ir
        self.lista_espect_muestras = file_paths_muestras_uv
        self.path_zero = file_path_zero_base_uv["ZeroLine"]
        self.path_base = file_path_zero_base_uv["BaseLine"]
        self.col_ir_ref = referencias_ir
        self.col_uv_ref = referencias_uv
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

    def leer_datos_referencia(self, column, type):
        if type == "uv":
            col_nm = "wvl [nm]"
        elif type == "ir":
            col_nm = "wvl [µm]"
        data_ref = pd.read_excel("references.xlsx", usecols=[col_nm, column])
        data_ref.rename(columns={col_nm: "nm"}, inplace=True)
        if col_nm == "wvl [µm]":
            data_ref["nm"] = data_ref["nm"] * 1000

        columnas_fijas = ['nm', 'Iz', 'Ib', 'I1', 'I2', 'I3']

        columna_a_renombrar = [col for col in data_ref.columns if col not in columnas_fijas][
            0]  # Obtiene la columna a renombrar

        data_ref.rename(columns={columna_a_renombrar: 'R_ref'}, inplace=True)
        return data_ref
    def leer_datos_muestra_IR(self):
        print(f"Leyendo datos IR de muestra {self.nombre}")
        ## Datos Tfir
        if self.tipo == "Absorbedor":
            tfir_base = pd.read_excel(self.archivo_tfir['archivo'], sheet_name=self.archivo_tfir["baseline"], skiprows=[0, 1, 2, 3])
            tfir_zero = pd.read_excel(self.archivo_tfir['archivo'], sheet_name=self.archivo_tfir["zeroline"], skiprows=[0, 1, 2, 3])

            tfir_base.rename(columns = {"%R": "Ib"}, inplace = True)
            tfir_zero.rename(columns={"%R": "Iz"}, inplace=True)
            # unir dataframes del ftir
            data_base_zero = pd.merge(tfir_zero, tfir_base, on="nm", how="inner")

            # Iterar sobre la lista de hojas y leer cada una
            i = 1
            for hoja in self.archivo_tfir["muestras"]:
                df = pd.read_excel(self.archivo_tfir['archivo'], sheet_name=hoja, skiprows=[0, 1, 2, 3])
                df.rename(columns = {"%R": f"I{i}"}, inplace = True)
                if i == 1:
                    data_ir = pd.merge(data_base_zero, df, on="nm", how="inner")
                else:
                    data_ir = pd.merge(data_ir, df, on="nm", how="inner")
                i = i + 1

            #leer datos de la referencia
            data_ref = self.leer_datos_referencia(self.col_ir_ref, "ir")

            data_ir = pd.merge(data_ir, data_ref, on="nm", how="inner")

            data_ir = self.calculos_medidas_ref_abs_col(data_ir)

        else:
            data_ir = np.nan
        return data_ir

    def leer_asc(selfself, path_asc):
        with open(path_asc, 'r') as file:
            lines = file.readlines()
        return lines
    def leer_datos_asc_columnas(self, path_asc, j, col_name):
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
        df = pd.DataFrame(data, columns=['nm', f"{col_name}{j}"]).astype(float)

        return df
    def leer_datos_muestra_UV(self):
        print(f"Leyendo datos UV de la muestra {self.nombre}")

        j=1
        data_uv =pd.DataFrame()
        lista_medidas = self.lista_espect_muestras
        for medida in lista_medidas:
            df = self.leer_datos_asc_columnas(medida, j, "I")
            if j == 1:
                data_uv = df
            else:
                data_uv = pd.merge(data_uv, df, on = "nm", how = "inner")
            j = j+1

        data_uv_base = self.leer_datos_asc_columnas(self.path_base, "b", "I")
        data_uv_zero = self.leer_datos_asc_columnas(self.path_zero, "z", "I")

        data_uv = pd.merge(data_uv, data_uv_base, on="nm", how="inner")
        data_uv = pd.merge(data_uv, data_uv_zero, on="nm", how="inner")

        # leer datos de la referencia
        data_ref = self.leer_datos_referencia(self.col_uv_ref, "uv")
        data_uv = pd.merge(data_uv, data_ref, on="nm", how="inner")

        data_uv.sort_values(by="nm", inplace=True)

        data_uv = self.calculos_medidas_ref_abs_col(data_uv)

        data_solar_w = self.leer_solar_w()
        data_uv = pd.merge(data_solar_w, data_uv, on = "nm", how = "inner")

        solar_w_ref, solar_w_abs = self.solar_w_ref_abd(data_uv)

        return data_uv, solar_w_ref, solar_w_abs

    def leer_solar_w(self):
        data_solar_w = pd.read_excel("references.xlsx", usecols=["wvl [nm]", "ASTM G173 direct"])
        data_solar_w.rename(columns={"wvl [nm]": "nm"}, inplace=True)
        return data_solar_w
    def combinar_uv_ir(self, data_ir, data_uv):
        if self.tipo == "Absorbedor":
            data_ir = data_ir.loc[data_ir["nm"]>2500]
            df_concatenado = pd.concat([data_ir, data_uv],
                                       ignore_index=True)  # df_uv antes para que quede de menor a mayor

            # Ordenar por 'nm'
            df_concatenado = df_concatenado.sort_values(by='nm')
        else:
            df_concatenado = data_uv

        return df_concatenado

    def calculos_medidas_ref_abs_col(self, df):
        # Identificar las columnas de Intensidad dinámicamente usando una expresión regular
        intensity_cols = [col for col in df.columns if re.match(r'^I\d+$', col)]

        # Calcular la media solo si existen columnas de intensidad
        if intensity_cols:
            df['Iw'] = df[intensity_cols].mean(axis=1)
        else:
            df['Iw'] = np.nan  # O algún valor por defecto si no hay columnas I

        df["ρw"] = (df["Iw"]-df["Iz"])/(df["Iw"]-df["Ib"])*df["R_ref"]
        df["αh"] = 1-df["ρw"]

        return df

    def solar_w_ref_abd(self, df):

        solar_w_ref = (df["ρw"] * df["ASTM G173 direct"]).sum()
        solar_w_abs = (df["αh"] * df["ASTM G173 direct"]).sum()

        return solar_w_ref, solar_w_abs

    def emitancia(self, df, temp):
        df["µm"] = df["nm"]/1000

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

        df["integrando"] = df["M_bb"] * df["αh"] * df["µm"].diff() * 1e-6

        df["denominador"] = df["M_bb"] * df["µm"].diff() * 1e-6

        emitancia = df["integrando"].sum() / df["denominador"].sum()

        return emitancia
