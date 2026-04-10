# Aliquindoi

*Aliquindoi*: voz malagueña que significa "estar atento". El nombre describe lo que hacían los técnicos antes de este programa.

Aliquindoi automatiza el procesado de datos de caracterización óptica de materiales solares. Lo que antes requería copiar manualmente mediciones entre instrumentos y plantillas durante horas, ahora se ejecuta en menos de un minuto.

Desarrollado por iniciativa propia para los técnicos del laboratorio OPAC de la Plataforma Solar de Almería, en el marco de la colaboración DLR-CIEMAT. Autora: Laura Campos Guzmán.

## Qué hace

El programa toma datos crudos de dos instrumentos de medición óptica — espectrofotómetro UV-VIS y espectrómetro FTIR — y genera informes estandarizados con los parámetros ópticos calculados para cuatro tipos de medida:

- **Reflectancia** — espejos solares CSP (SWR ponderado con E891 e ISO 9050)
- **Absortancia** — absorbedores solares selectivos (SWA solar + emitancia térmica)
- **Transmitancia CSP** — vidrios de cubierta para colectores
- **Transmitancia PV** — vidrios para módulos fotovoltaicos

Los cálculos siguen los estándares ASTM G173 y E891. La emitancia térmica se obtiene por integración numérica de la ley de Planck sobre el espectro IR medido.

## Demo

Esta versión pública sustituye la GUI interactiva y las plantillas propietarias del entorno de laboratorio por una demo autocontenida con datos sintéticos:

```bash
pip install -r requirements_demo.txt
python main_demo.py
```

Genera en `demo_output/`:

- `resultados_demo.csv` — tabla con SWR, SWA, transmitancia solar y emitancia para las 4 medidas
- `espectros_demo.html` — gráficos interactivos de los espectros (abrir en navegador)

## Estructura

```
Aliquindoi/
├── muestra.py              # Objeto de dominio: cálculos ópticos (numpy/pandas)
├── programas/
│   ├── lectura_datos.py    # I/O y diálogos GUI (entorno de laboratorio)
│   └── plantillas_excel.py # Salida a plantillas Excel (entorno de laboratorio)
main_demo.py                # Demo pública: sin GUI, sin Excel, datos sintéticos
demo_data/                  # Datos sintéticos y espectros solares de referencia
requirements_demo.txt       # Dependencias mínimas para el demo
```

## Entorno de laboratorio vs. demo pública

El programa en producción usa `xlwings` para escribir directamente en plantillas Excel con formato estandarizado DLR/CIEMAT, y una GUI Tkinter para que los técnicos seleccionen archivos y parámetros de medición. Esos componentes dependen de Windows + Excel instalado y de plantillas propietarias que no se pueden distribuir.

El motor de cálculo (`muestra.py`) es independiente de la GUI y del formato de salida, y es el que se usa en esta demo.

### Código del programa original

La carpeta `Aliquindoi/` contiene el código completo del programa en producción:

- `muestra.py` — motor de cálculo: objeto de dominio con toda la lógica óptica (lectura de espectros UV e IR, cálculo de reflectancia/absortancia/transmitancia, integración numérica de la ley de Planck para la emitancia térmica)
- `main_aliquindoi.py` — orquestador principal: setup, configuración de referencias por instrumento, loop de muestras, manejo de errores y limpieza de recursos
- `programas/lectura_datos.py` — capa de entrada: GUI Tkinter, diálogos de selección de archivos, búsqueda automática de muestras y referencias en el sistema de ficheros
- `programas/plantillas_excel.py` — capa de salida: escritura en plantillas Excel DLR/CIEMAT vía xlwings

Este código no es ejecutable fuera del entorno de laboratorio (requiere Windows + Excel + plantillas propietarias), pero refleja la arquitectura real del sistema y las decisiones de diseño tomadas en producción.

## Dependencias

```
numpy
pandas
plotly
openpyxl
```
