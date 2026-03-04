# Arquitectura actual de Aliquindoi

## Introducción

Este documento describe la arquitectura y el flujo de ejecución del programa Aliquindoi **tal como está implementado actualmente**. No se proponen mejoras ni refactorizaciones; el objetivo es únicamente documentar el comportamiento existente para facilitar su comprensión.

El programa procesa datos espectrales de mediciones ópticas provenientes de equipos FTIR (infrarrojo) y espectrofotómetros UV, realizando cálculos de reflectancia, absortancia y emitancia térmica, y generando reportes en Excel basados en plantillas configurables.

---

## Punto de entrada

**Archivo:** [`main_aliquindoi.py`](file:///d:/Archivos/opac/OPAC_programa/OPAC/Aliquindoi/main_aliquindoi.py)

El programa inicia su ejecución en este archivo, que actúa como orquestador principal. La secuencia de inicio es:

1. Carga la configuración desde [`importar_configuracion.py`](file:///d:/Archivos/opac/OPAC_programa/OPAC/user_templates/Configuracion/importar_configuracion.py) (lee `config.ini`)
2. Solicita al usuario la ubicación del archivo Excel de salida mediante un diálogo
3. Crea un libro Excel vacío en la ubicación especificada
4. Abre el libro con `xlwings` y desactiva actualizaciones de pantalla para mejorar el rendimiento

---

## Flujo de ejecución de alto nivel

### 1. Recopilación de datos del usuario

**Función:** `lectura_datos.pregunta_tipos_test()`

Se muestra una interfaz gráfica (Tkinter) que recopila:
- **Tipo de medida:** Reflectancia, Absortancia, Transmitancia CSP o Transmitancia PV
- **Equipos utilizados:** FTIR y/o Espectrofotómetro (con o sin ventana)
- **Metadatos del test:** Sitio de test, fabricante, proyecto
- **Parámetros temporales:** Horas de exposición, meses, temperatura
- **Fecha de medida:** Seleccionada mediante widget de calendario

**Resultado:** Diccionario `datos_basicos` con todas las selecciones del usuario

### 2. Selección de rutas de archivos

Según los equipos seleccionados:

**Si se selecciona FTIR:**
- Usuario selecciona carpeta con datos FTIR
- Se determina si usa ventana (`ventana_ftir = True/False`)
- Usuario selecciona columnas de referencia desde `references.xlsx`:
  - Referencia FTIR (`r_ftir`)
  - Transmitancia de ventana si aplica (`r_trans_ir`)

**Si se selecciona Espectrofotómetro:**
- Usuario selecciona carpeta con datos del espectrofotómetro
- Se determina si usa ventana (`ventana_esp = True/False`)
- Usuario selecciona columnas de referencia según el tipo de medida:
  - Para Absortancia: referencia de absorbancia
  - Para Reflectancia: referencia de reflectancia
  - Para Transmitancia CSP: referencia de transmitancia
  - Transmitancia de ventana si aplica (`r_trans_uv`)

### 3. Detección automática de nombres de muestras

**Función:** `lectura_datos.nombres_muestras_auto()`

Extrae automáticamente los nombres de las muestras desde:
- **Ruta espectrofotómetro:** Busca archivos `.asc` que empiecen con `sample_` y extrae nombres usando patrón regex `sample_(.*)-[^-]+`
- **Ruta FTIR** (si no hay espectrofotómetro): Busca archivos Excel con "data" en el nombre y extrae nombres desde hojas que empiecen con `sample_`

**Resultado:** Lista de nombres únicos de muestras (ordenados, sin duplicados)

### 4. Procesamiento de cada muestra (bucle principal)

Para cada nombre de muestra en `nombres_muestras`:

#### 4.1 Descubrimiento de archivos

**Archivos FTIR** (si FTIR seleccionado):
- **Función:** `lectura_datos.ftir_medidas_auto()`
  - Busca archivos Excel con "data" en el nombre
  - Encuentra hojas que coincidan con `sample_<nombre>`
  - Extrae fecha de medida del nombre del archivo
  - Localiza archivo de referencia que contenga fecha + "ref"
  - Identifica hojas de referencia: `zero_`, `base_`, `ventana_` (si hay ventana)

**Archivos UV Espectrofotómetro** (si Espectrofotómetro seleccionado):
- **Función:** `lectura_datos.espectro_medidas_zero_base_auto()`
  - Encuentra archivos `.asc` de la muestra (`sample_<nombre>-*.asc`)
  - Para Reflectancia/Absortancia, también encuentra:
    - Archivos de línea cero (`zero_*.asc`)
    - Archivos de línea base (`base_*.asc`)
    - Archivos de ventana si aplica (`ventana_*.asc`)
  - Si existen múltiples archivos de línea base, selecciona el más cercano en tiempo a la medida de la muestra

#### 4.2 Creación de instancia de muestra

**Clase:** `Muestra` ([`muestra.py`](file:///d:/Archivos/opac/OPAC_programa/OPAC/Aliquindoi/muestra.py))

Se instancia un objeto `Muestra` que almacena:
- Nombre de la muestra, rutas a todos los archivos de datos
- Selecciones de columnas de referencia
- Tipo de medida y configuración de equipos
- Metadatos (test, fabricante, proyecto, horas, meses, temperatura, fecha)

#### 4.3 Procesamiento de datos (ramificación por tipo de medida)

##### **Rama A: Reflectancia**

1. Llama a `plantillas_excel.copiar_datos_excel_config()`
2. La muestra procesa datos UV:
   - Lee archivos `.asc` (`leer_datos_UV()`)
   - No realiza cálculos adicionales
3. Copia datos directamente a plantilla Excel
4. Plantilla seleccionada según rango de longitud de onda (280nm, 300nm o 320nm)

##### **Rama B: Absortancia**

1. **Procesamiento FTIR** (si existen datos FTIR):
   - `Muestra.procesar_datos_tfir()`: Lee hojas Excel, combina datos de muestra y referencia
   - `Muestra.medidas_ir()`: Calcula reflectancia y absortancia IR
     - Fórmula (sin ventana): `refl = (Iw - zero) / (base - zero) * r_ftir`
     - Fórmula (con ventana): Corrige efectos de ventana
     - `abs = 1 - refl`

2. **Procesamiento UV** (si existen datos de Espectrofotómetro):
   - `Muestra.leer_datos_UV()`: Lee archivos `.asc`, fusiona con referencias y espectro ASTM
   - `Muestra.medidas_UV()`: Calcula reflectancia y absortancia UV
     - Promedia múltiples medidas de intensidad
     - Fórmula (sin ventana): `refl = ((Iw - zero) / (base - zero)) * r_uv`
     - Fórmula (con ventana): Corrige efectos de ventana
     - `abs = 1 - refl`
     - Calcula reflectancia ponderada solar (SWR) y absortancia (SWA)
     - Calcula desviación estándar de SWR

3. **Combinar UV e IR**:
   - `Muestra.combinar_uv_ir()`: Concatena datos IR (2500-16000 nm) y UV

4. **Cálculo de emitancia** (si existen datos FTIR y UV):
   - `Muestra.emitancia()`: Calcula emitancia térmica usando ley de Planck
   - Integra `M_bb * abs * d(lambda)` sobre rango de longitudes de onda
   - Usa temperatura especificada por usuario

5. **Escritura a Excel**:
   - `plantillas_excel.copiar_datos_excel_absorbedores_config()`
   - Escribe resultados calculados y datos en bruto a plantilla

##### **Rama C: Transmitancia (CSP o PV)**

1. Procesa datos UV sin correcciones de línea base (no hay archivos `zero` o `base`)
2. Llama a `plantillas_excel.copiar_datos_excel()` o `copiar_datos_excel_config()`
3. Copia datos de transmitancia a plantilla apropiada

### 5. Población de plantillas

**Módulo:** `plantillas_excel.py`

Para cada muestra:

1. **Selección de plantilla:**
   - Lee longitud de onda de la última línea del archivo `.asc`
   - Selecciona plantilla apropiada según:
     - Tipo de medida (Reflectancia, Absortancia, Transmitancia CSP/PV)
     - Rango de longitud de onda (280nm, 300nm, 320nm)
   - Rutas de plantillas leídas desde `config.ini`

2. **Inserción de datos:**
   - Abre plantilla Excel usando `xlwings`
   - Inserta metadatos (nombre muestra, fabricante, tipo test, fecha, etc.)
   - Inserta datos de medición:
     - Valores de intensidad en bruto de archivos `.asc`
     - Valores calculados de reflectancia/absortancia
     - Datos de referencia
   - Elimina columnas de medición no usadas según número real de mediciones

3. **Creación de hoja:**
   - Copia hoja de plantilla poblada al libro de salida
   - Renombra hoja como `<nombre_muestra>_<fecha_id>` (truncado a 31 caracteres)
   - Cierra archivo de plantilla

---

## Módulos principales y sus responsabilidades

### [`main_aliquindoi.py`](file:///d:/Archivos/opac/OPAC_programa/OPAC/Aliquindoi/main_aliquindoi.py)
- **Rol:** Orquestador principal y punto de entrada
- **Responsabilidades:**
  - Inicializa configuración y libro de salida
  - Recopila entradas del usuario
  - Itera sobre muestras
  - Despacha a rama de procesamiento apropiada
  - Manejo de errores y logging

### [`lectura_datos.py`](file:///d:/Archivos/opac/OPAC_programa/OPAC/Aliquindoi/programas/lectura_datos.py)
- **Rol:** Interfaz de usuario y descubrimiento de archivos
- **Responsabilidades:**
  - Todos los diálogos GUI Tkinter para entrada de usuario
  - Diálogos de selección de archivos y carpetas
  - Extracción automática de nombres de muestras
  - Resolución de rutas de archivos (Excel FTIR, `.asc` UV)
  - Selección de referencias desde `references.xlsx`

### [`muestra.py`](file:///d:/Archivos/opac/OPAC_programa/OPAC/Aliquindoi/muestra.py)
- **Rol:** Modelo de datos y motor de cálculo
- **Responsabilidades:**
  - Representa una muestra individual con todas sus mediciones
  - Lee datos en bruto de archivos (Excel FTIR, `.asc` UV)
  - Realiza cálculos ópticos (reflectancia, absortancia, emitancia)
  - Fusiona datos de múltiples fuentes
  - Lee valores de referencia desde `references.xlsx`

### [`plantillas_excel.py`](file:///d:/Archivos/opac/OPAC_programa/OPAC/Aliquindoi/programas/plantillas_excel.py)
- **Rol:** Gestión de plantillas Excel y generación de salida
- **Responsabilidades:**
  - Selecciona plantilla Excel apropiada según tipo de medida y longitud de onda
  - Parsea archivos `.asc` para extraer valores de medición
  - Puebla plantillas con datos y metadatos
  - Gestiona mapeos de celdas desde configuración
  - Crea hojas en libro de salida

### [`importar_configuracion.py`](file:///d:/Archivos/opac/OPAC_programa/OPAC/user_templates/Configuracion/importar_configuracion.py)
- **Rol:** Cargador de configuración
- **Responsabilidades:**
  - Lee archivo `config.ini`
  - Provee rutas de archivos de plantilla
  - Provee configuraciones de mapeo de celdas para diferentes tipos de plantilla

---

## Puntos de decisión y ramificaciones

### Decisión 1: Selección de equipos
- **Ubicación:** Líneas 34-80 en `main_aliquindoi.py`
- **Condición:** Basada en `datos_basicos["aparatos"]`
- **Ramas:**
  - Solo FTIR → Procesa solo datos IR
  - Solo Espectrofotómetro → Procesa solo datos UV
  - Ambos → Procesa y combina datos IR + UV

### Decisión 2: Uso de ventana
- **Ubicación:** Líneas 37-40 (FTIR), 56-59 (UV) en `main_aliquindoi.py`
- **Condición:** Selección de usuario "Con ventana" vs "Sin ventana"
- **Impacto:** 
  - Cambia requisitos de archivos de referencia
  - Modifica fórmulas de cálculo en `medidas_UV()` y `medidas_ir()`

### Decisión 3: Tipo de medida
- **Ubicación:** Líneas 109-140 en `main_aliquindoi.py`
- **Condición:** `datos_basicos["medida"]`
- **Ramas:**
  - **Reflectancia:** Copia directa a plantilla, procesamiento mínimo
  - **Absortancia:** Pipeline completo de cálculo, cálculo de emitancia
  - **Transmitancia CSP/PV:** Plantilla diferente, sin archivos de línea base

### Decisión 4: Cálculo de emitancia
- **Ubicación:** Líneas 129-133 en `main_aliquindoi.py`
- **Condición:** Datos de FTIR y Espectrofotómetro disponibles
- **Ramas:**
  - Ambos presentes → Calcula emitancia
  - Uno faltante → Establece emitancia a "no calculada"

### Decisión 5: Selección de plantilla
- **Ubicación:** `plantillas_excel.elegir_plantilla_config()` (líneas 113-187)
- **Condición:** Tipo de medida + rango de longitud de onda (de última línea de archivo `.asc`)
- **Ramas:**
  - 9 combinaciones diferentes (3 tipos de medida × 3 rangos de longitud de onda + PV)
  - Cada una usa diferente plantilla Excel y mapeos de celdas

### Decisión 6: Selección de archivo de línea base
- **Ubicación:** `lectura_datos.espectro_medidas_zero_base_auto()` (líneas 456-478)
- **Condición:** Múltiples archivos de línea base encontrados
- **Lógica:** 
  - Lee timestamps de archivos de línea base y muestra
  - Selecciona archivo de línea base con timestamp más cercano al tiempo de medida de la muestra

---

## Entradas, salidas y efectos secundarios

### Entradas

#### Del usuario (diálogos GUI):
1. Ruta del archivo Excel de salida
2. Metadatos de medición:
   - Tipo (Reflectancia/Absortancia/Transmitancia)
   - Equipo (FTIR/Espectrofotómetro, con/sin ventana)
   - Sitio de test, fabricante, proyecto
   - Horas, meses, temperatura
   - Fecha de medida
3. Rutas de directorios:
   - Carpeta de datos FTIR
   - Carpeta de datos espectrofotómetro
4. Selecciones de referencia desde `references.xlsx`:
   - Columna de referencia UV (`r_uv`)
   - Columna de referencia FTIR (`r_ftir`)
   - Columnas de transmitancia de ventana (si aplica)

#### Del sistema de archivos:
1. **Archivo de configuración:** `config.ini` (rutas de plantillas y mapeos de celdas)
2. **Datos de referencia:** `references.xlsx` (propiedades ópticas de materiales de referencia, espectro ASTM)
3. **Datos FTIR:** Archivos Excel (`.xlsx`) con patrón de nombre `*data*.xlsx`
   - Contiene hojas: `sample_<nombre>`, `zero_`, `base_`, `ventana_` (opcional)
4. **Datos UV:** Archivos ASCII (`.asc`) con patrones de nombre:
   - Muestra: `sample_<nombre>-*.Sample.asc`
   - Línea base: `base_*.asc`
   - Cero: `zero_*.asc`
   - Ventana: `ventana_*.asc` (opcional)
5. **Plantillas Excel:** Ubicadas en `user_templates/plantillas/`
   - Múltiples plantillas para diferentes tipos de medida y rangos de longitud de onda

### Salidas

#### Libro Excel:
- **Ubicación:** Ruta especificada por usuario (del diálogo inicial)
- **Estructura:**
  - Una hoja por muestra
  - Nombres de hoja: `<nombre_muestra>_<YYYYMMDD>` (máx 31 caracteres)
  - Cada hoja contiene:
    - Metadatos (información de muestra, parámetros de test)
    - Datos de medición (longitud de onda, reflectancia, absortancia)
    - Valores calculados (SWR, SWA, emitancia para medidas de absortancia)
    - Tablas de datos en bruto (mediciones originales FTIR/UV)

### Efectos secundarios

#### Archivos escritos:
1. **Salida principal:** Libro Excel en ubicación especificada por usuario
2. **Log de errores** (en caso de excepción): Archivo `.txt` en mismo directorio que Excel de salida
   - Creado reemplazando extensión `.xlsx` por `.txt`
   - Contiene timestamp y traceback completo

#### Archivos leídos:
- Todos los archivos de entrada listados arriba (acceso de solo lectura)

#### Estado de aplicación Excel:
- Durante ejecución:
  - Actualización de pantalla desactivada
  - Cálculo automático desactivado
  - Alertas de visualización desactivadas
- Al completar (en bloque `finally`):
  - Actualización de pantalla reactivada
  - Cálculo automático reactivado
  - Libro de salida guardado y cerrado

#### Salida de consola:
- Mensajes de progreso impresos durante toda la ejecución
- Rutas de archivos siendo procesados
- Nombres de muestras detectados
- Errores y advertencias
- Confirmaciones de procesamiento de datos

---

## Algoritmos clave

### Cálculo de reflectancia (sin ventana)
```
refl = ((Intensidad_muestra - Cero) / (Linea_base - Cero)) * Reflectancia_referencia
```

### Cálculo de reflectancia (con ventana)
```
refl_ventana = ((Intensidad_ventana - Cero) / (Linea_base - Cero)) * Referencia
refl_muestra = ((Intensidad_muestra - Cero) / (Linea_base - Cero)) * Referencia
refl_corregida = (refl_muestra - refl_ventana) / 
                 ((Transmitancia_ventana/100)² + refl_ventana * (refl_muestra - refl_ventana))
```

### Cálculo de absortancia
```
abs = 1 - refl
```

### Reflectancia ponderada solar (SWR)
```
SWR = Σ(refl * espectro_ASTM) / Σ(espectro_ASTM)
```

### Emitancia térmica
```
M_bb(λ,T) = (2πhc²/λ⁵) * 1/(exp(hc/λkT) - 1)  [Ley de Planck]
emitancia = Σ(M_bb * abs * dλ) / Σ(M_bb * dλ)
```

---

## Manejo de errores

- **Bloque try-except principal:** Envuelve toda la ejecución (líneas 14-141 en `main_aliquindoi.py`)
- **En caso de error:**
  - Captura traceback completo
  - Escribe a archivo de log (`.txt` con timestamp)
  - Imprime ubicación del error en consola
- **Bloque finally:** Asegura limpieza de Excel (restaurar configuración, guardar, cerrar)
- **Errores de funciones individuales:** 
  - Errores de lectura de archivos imprimen advertencias pero continúan
  - Datos faltantes resultan en DataFrames vacíos o valores "no calculada"

---

## Sistema de configuración

El programa usa configuración de dos niveles:

1. **`config.ini`:** Define rutas de plantillas y mapeos de celdas para cada tipo de medida
2. **Clase `Config`:** Carga configuración al inicio, provee acceso a:
   - `path_plantillas`: Diccionario de rutas de archivos de plantilla
   - `celdas_plantillas`: Diccionario de mapeos de celdas para cada tipo de plantilla

Esto permite cambiar plantillas y ubicaciones de celdas sin modificar código.

---

## Tipos de plantilla

El programa soporta estas combinaciones de plantillas:

| Tipo de medida | Longitud de onda | Clave de plantilla |
|----------------|------------------|--------------------|
| Reflectancia | 280nm | `refl_280` |
| Reflectancia | 300nm | `refl_300` |
| Reflectancia | 320nm | `refl_320` |
| Absortancia | 280nm | `abs_280` |
| Absortancia | 300nm | `abs_300` |
| Absortancia | 320nm | `abs_320` |
| Transmitancia CSP | 280nm | `trans_csp_280` |
| Transmitancia CSP | 300nm | `trans_csp_300` |
| Transmitancia CSP | 320nm | `trans_csp_320` |
| Transmitancia PV | N/A | `trans_pv` |

Cada tipo de plantilla tiene su propia estructura de hoja y configuración de mapeo de celdas.

---

## Flujo de datos resumido

```
Entrada de usuario (GUI)
    ↓
Descubrimiento de archivos (lectura_datos)
    ↓
Para cada muestra:
    ↓
    Resolución de rutas → Creación de instancia Muestra
    ↓                           ↓
    Lectura de datos ← ← ← ← ← ←
    (archivos .asc sin procesar, hojas Excel)
    ↓
    Cálculos ópticos
    (medidas_UV, medidas_ir, combinar_uv_ir, emitancia)
    ↓
    Selección de plantilla
    (según tipo de medida + longitud de onda)
    ↓
    Población de Excel
    (copiar plantilla → insertar datos → renombrar hoja)
    ↓
Guardar libro de salida
```

---

## Notas finales

Esta documentación refleja el estado actual del programa sin proponer cambios. El objetivo es servir como referencia para entender el flujo de ejecución, las decisiones de diseño implementadas y la estructura modular existente.
