from enum import unique
from unittest.mock import inplace

import pandas as pd
import os
import re

path_espectofotometro = "C:/Users/camp_la/Documents/Laura/OPAC/proyecto_automatizar_procesos/reflectores/1000h/20241216_l2"

def encontrar_archivos_asc_por_nombre(nombre, carpeta):
    archivos_encontrados = []
    for raiz, directorios, archivos in os.walk(carpeta):
        for archivo in archivos:
            if archivo.startswith(nombre) and archivo.endswith(".asc"):
                ruta_completa = os.path.join(raiz, archivo)
                archivos_encontrados.append(ruta_completa)
    return archivos_encontrados

archivos = encontrar_archivos_asc_por_nombre("sample_", path_espectofotometro)
print(archivos)
resultados = [re.search(r'sample_(.*)-[^-]+$', archivo).group(1) for archivo in archivos if re.search(r'sample_(.*)-[^-]+\.Sample\.asc$', archivo)]
print(resultados)
muestras = list(set(resultados))
print(muestras)


