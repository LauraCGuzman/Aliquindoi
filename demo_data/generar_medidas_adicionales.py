"""
Genera datos sintéticos para Reflectancia, Transmitancia CSP y Transmitancia PV.

Estructura generada:
  demo_data/
    refl/
      sample_ReflCSP-1.Sample.asc   (reflector solar ~0.93 visible, caida en IR)
      sample_ReflCSP-2.Sample.asc
      zero_refl.asc
      base_refl.asc
    trans_csp/
      sample_GlassCSP-1.Sample.asc  (vidrio CSP: T~0.91, absorcion UV < 320 nm)
      sample_GlassCSP-2.Sample.asc
    trans_pv/
      sample_GlassPV-1.Sample.asc   (vidrio PV: T alta visible, corte UV a 380 nm)
      sample_GlassPV-2.Sample.asc

Formato .asc identico al generado para UV/absortancia:
  linea 0 : titulo
  linea 1-2: vacias
  linea 3 : fecha yy/mm/dd
  linea 4 : hora HH:MM:SS.ffffff
  lineas 5-6: cabecera libre
  #DATA
  nm<TAB>valor  (coma como separador decimal)
"""

import os
import numpy as np

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FECHA_STR = "26/04/09"
np.random.seed(7)

NM = np.arange(250, 2505, 5)   # 250-2500 nm, igual que UV/absortancia


# ── utilidad ──────────────────────────────────────────────────────────────────

def escribir_asc(ruta, nm_array, valores, hora, titulo):
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(f"{titulo}\n")
        f.write("\n")
        f.write("\n")
        f.write(f"{FECHA_STR}\n")
        f.write(f"{hora}\n")
        f.write("Unidades: nm / %\n")
        f.write("Instrumento: Demo Spectrophotometer\n")
        f.write("#DATA\n")
        for nm_val, val in zip(nm_array, valores):
            f.write(f"{str(nm_val).replace('.', ',')}\t{str(round(float(val), 6)).replace('.', ',')}\n")
    print(f"  OK  {os.path.relpath(ruta, BASE_DIR)}")


def sigmoide(nm, centro, anchura, invertida=False):
    """Transicion sigmoide suave, util para cortes UV."""
    s = 1.0 / (1.0 + np.exp(-(nm - centro) / anchura))
    return (1.0 - s) if invertida else s


# ── perfiles espectrales ──────────────────────────────────────────────────────

def perfil_reflector_csp(nm, replica=0):
    """
    Reflector solar CSP (espejo de plata/aluminio):
    - Reflectancia ~0.93 plana en visible (400-800 nm)
    - Caida gradual en NIR (800-2500 nm) hasta ~0.88
    - Absorcion UV leve por debajo de 380 nm
    - Ruido instrumental pequeño entre replicas
    """
    caida = 0.05 * (np.maximum(0, nm - 400) / 2100) ** 1.5
    base = np.where(nm < 400,
                    0.85 + 0.08 * sigmoide(nm, 360, 20),
                    0.93 - caida)
    base += np.random.normal(0, 0.003 + replica * 0.001, size=len(nm))
    return np.clip(base, 0.01, 0.99)


def perfil_zero(nm):
    return np.clip(np.random.normal(0.001, 0.0005, size=len(nm)), 0.0, 0.01)


def perfil_base(nm):
    return np.clip(0.99 + np.random.normal(0, 0.002, size=len(nm)), 0.92, 1.02)


def perfil_vidrio_csp(nm, replica=0):
    """
    Vidrio solar para CSP (borosilicato de baja reflectancia):
    - Transmitancia ~0.91 en visible (380-1200 nm)
    - Corte UV gradual por debajo de 320 nm (absorcion por Fe2O3 bajo)
    - Caida en NIR a partir de 1200 nm (absorcion OH y agua estructural)
    - Ruido pequeno entre replicas
    """
    corte_uv = sigmoide(nm, 310, 15)          # sube de 0 a 1 alrededor de 310 nm
    caida_nir = 1.0 - 0.15 * sigmoide(nm, 1400, 200)  # caida suave a partir de 1400 nm
    base = 0.916 * corte_uv * caida_nir
    base -= 0.01 * np.exp(-((nm - 950) ** 2) / (2 * 150 ** 2))  # banda de absorcion ~950 nm
    base += np.random.normal(0, 0.003 + replica * 0.001, size=len(nm))
    return np.clip(base, 0.0, 0.99)


def perfil_vidrio_pv(nm, replica=0):
    """
    Vidrio solar para PV (vidrio extra-claro, corte UV a 380 nm):
    - Corte brusco por debajo de 380 nm (sin transmision UV)
    - Transmitancia muy alta en visible: ~0.92 (400-1100 nm)
    - Caida en NIR a partir de 1100 nm
    - La banda UV actua como filtro para proteger las celulas
    """
    corte_uv = sigmoide(nm, 375, 8)           # transicion rapida en 375 nm
    caida_nir = 1.0 - 0.25 * sigmoide(nm, 1200, 150)
    base = 0.925 * corte_uv * caida_nir
    base -= 0.008 * np.exp(-((nm - 1400) ** 2) / (2 * 200 ** 2))  # banda ~1400 nm
    base += np.random.normal(0, 0.003 + replica * 0.001, size=len(nm))
    return np.clip(base, 0.0, 0.99)


# ── generacion ────────────────────────────────────────────────────────────────

def generar_reflectancia():
    print("\n=== demo_data/refl/ (Reflectancia) ===")
    d = os.path.join(BASE_DIR, "refl")

    escribir_asc(os.path.join(d, "zero_refl.asc"),
                 NM, perfil_zero(NM), "08:55:00.000000", "ZeroLine")
    escribir_asc(os.path.join(d, "base_refl.asc"),
                 NM, perfil_base(NM), "09:00:00.000000", "BaseLine")

    for i in range(1, 3):
        hora = f"09:{i * 15:02d}:00.000000"
        escribir_asc(os.path.join(d, f"sample_ReflCSP-{i}.Sample.asc"),
                     NM, perfil_reflector_csp(NM, replica=i - 1),
                     hora, f"ReflCSP replica {i}")


def generar_trans_csp():
    print("\n=== demo_data/trans_csp/ (Transmitancia CSP) ===")
    d = os.path.join(BASE_DIR, "trans_csp")

    for i in range(1, 3):
        hora = f"10:{i * 15:02d}:00.000000"
        escribir_asc(os.path.join(d, f"sample_GlassCSP-{i}.Sample.asc"),
                     NM, perfil_vidrio_csp(NM, replica=i - 1),
                     hora, f"GlassCSP replica {i}")


def generar_trans_pv():
    print("\n=== demo_data/trans_pv/ (Transmitancia PV) ===")
    d = os.path.join(BASE_DIR, "trans_pv")

    for i in range(1, 3):
        hora = f"11:{i * 15:02d}:00.000000"
        escribir_asc(os.path.join(d, f"sample_GlassPV-{i}.Sample.asc"),
                     NM, perfil_vidrio_pv(NM, replica=i - 1),
                     hora, f"GlassPV replica {i}")


if __name__ == "__main__":
    generar_reflectancia()
    generar_trans_csp()
    generar_trans_pv()
    print("\nDone.")
