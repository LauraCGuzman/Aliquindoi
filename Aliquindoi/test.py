file = "C:/Users/camp_la/Documents/Laura/OPAC/proyects/2023_ADVIAMOS/800º_mufla/1000h_800º_mufla/20231127/base_prueba.asc"

with open(file, 'r') as archivo:
    lineas = archivo.readlines()  # Lee todas las líneas en una lista

    # Imprime las líneas 3 y 4 (índices 2 y 3 en Python)
    for i in range(3, 5):
        if (i == 3):
            fecha = lineas[i]
        elif (i==4):
            hora = lineas[i]
