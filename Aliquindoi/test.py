import pandas as pd
path_asc = "C:/Users/camp_la/Documents/Laura/OPAC/proyecto_automatizar_procesos/reflectores/1000h/20241216 L1/51-1-1.Sample.asc"

with open(path_asc, 'r') as file:
    lines = file.readlines()

del lines[73:79]

i = 1
for line in lines:
    print(i, ": ", line)
    i = i+1
