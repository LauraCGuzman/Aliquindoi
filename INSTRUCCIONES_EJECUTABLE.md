# 🚀 Guía: Creación del Ejecutable (Aliquindoi)

Este documento detalla los pasos para generar el archivo `.exe` de forma limpia.

## 1. Instalación de Herramientas

🚀 Guía: Creación del Ejecutable (Aliquindoi)
Este documento detalla los pasos para generar el archivo .exe de forma limpia.

1. Instalación de Herramientas
Ejecuta estos comandos en la terminal de PyCharm para asegurar que tienes todo lo necesario:
```bash
pip install -r requirements.txt
pip install pyinstaller
```

## 2. Comando para Generar el Ejecutable
Copia y pega el siguiente bloque completo en tu terminal (PowerShell o Git Bash). La barra \ permite que el comando se lea como una sola línea aunque esté separado para mayor claridad:
```bash
pyinstaller --noconfirm --clean --name Aliquindoi \
--add-data "user_templates;user_templates" \
--add-data "Aliquindoi;Aliquindoi" \
--paths "." \
--console Aliquindoi/main_aliquindoi.py
```
