
import sys
import os
from unittest.mock import MagicMock

# Mock xlwings and others to avoid import errors
sys.modules["xlwings"] = MagicMock()
sys.modules["openpyxl"] = MagicMock()
sys.modules["openpyxl.utils.cell"] = MagicMock()

# Now import
# Ensure we can import plantillas_excel from current dir
sys.path.append(os.getcwd())
from plantillas_excel import elegir_plantilla_config

class MockMuestra:
    tipo_medida = "Reflectancia"

class MockConfig:
    # Use the exact path from config.ini we saw earlier
    path_plantillas = {"reflectance": "../user_templates/plantillas/refl_plantilla.xls"}
    celdas_plantillas = {"reflectance": {}}

config = MockConfig()
muestra = MockMuestra()

print(f"Testing 'elegir_plantilla_config'...")
try:
    path, celdas = elegir_plantilla_config(muestra, config)
    print(f"Resolved Path: {path}")

    if os.path.exists(path):
        print("SUCCESS: File exists on disk.")
    else:
        print("FAILURE: File path returned but does not exist.")
except Exception as e:
    print(f"EXCEPTION: {e}")
