"""
main_demo.py — Pipeline completo de Aliquindoi con datos sintéticos.
Cubre los cuatro tipos de medida: Absortancia, Reflectancia, Transmitancia CSP, Transmitancia PV.
Replica exactamente los cálculos del programa real para cada tipo.
Genera demo_resultados.html con 4 subgráficos Plotly interactivos.
"""

import sys
import os

# Forzar UTF-8 en stdout para que los caracteres especiales no fallen en Windows
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from Aliquindoi.muestra import Muestra

# ── Rutas ──────────────────────────────────────────────────────────────────────
DEMO_DIR   = os.path.join(current_dir, "demo_data")
UV_DIR     = os.path.join(DEMO_DIR, "UV")
REFL_DIR   = os.path.join(DEMO_DIR, "refl")
TCSP_DIR   = os.path.join(DEMO_DIR, "trans_csp")
TPV_DIR    = os.path.join(DEMO_DIR, "trans_pv")
FTIR_DIR   = os.path.join(DEMO_DIR, "FTIR")
REFS_PATH  = os.path.join(current_dir, "user_templates", "references.xlsx")
SOLAR_PATH = os.path.join(DEMO_DIR, "solar_spectra.csv")
HTML_OUT   = os.path.join(current_dir, "demo_resultados.html")

FTIR_DATA = os.path.join(FTIR_DIR, "20260409data.xlsx")
FTIR_REF  = os.path.join(FTIR_DIR, "20260409ref.xlsx")

# ── Parámetros del ensayo de absortancia ───────────────────────────────────────
DATOS_BASICOS_ABS = {
    "medida":     "Absortancia",
    "aparatos":   {"FTIR": "Sin ventana", "Espectrofotómetro": "Sin ventana"},
    "test":       "Demo", "fabricante": "Demo", "proyecto": "Demo",
    "hours":      None, "months": None, "temperatura": 373.0,
    "fecha_medida": {"dd/mm/yyyy": "09/04/2026", "yyyyMMdd": "20260409"},
}
REFS_IR = {"r_ftir": "FS1 (specular FISE absorber)", "r_trans_ir": ""}
REFS_UV = {"r_uv": "Grey spectralon 10% (provided by Octalab in 2021)", "r_trans_uv": ""}

# ── Utilidades ──────────────────────────────────────────────────────────────────

def leer_asc(path, col_name):
    """Replica exacta de Muestra.leer_datos_asc_columnas (muestra.py:215)."""
    with open(path, 'r') as f:
        lines = f.readlines()
    data_start = next(i + 1 for i, l in enumerate(lines) if '#DATA' in l)
    data = [l.strip().replace(',', '.').split('\t') for l in lines[data_start:] if l.strip()]
    return pd.DataFrame(data, columns=['nm', col_name]).astype(float)


def leer_espectro_solar():
    return pd.read_csv(SOLAR_PATH)


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 1 — ABSORTANCIA
#  Flujo: procesar_datos_tfir → medidas_ir → leer_datos_UV → medidas_UV
#         → combinar_uv_ir → emitancia
#  (idéntico a main_aliquindoi.py:157-188)
# ═══════════════════════════════════════════════════════════════════════════════

print("=" * 60)
print("ALIQUINDOI — DEMO (pipeline completo)")
print("=" * 60)
print("\n── ABSORTANCIA ──")

MUESTRAS_ABS = [
    (
        "SampleA",
        ["sample_SampleA-1", "sample_SampleA-2"],
        [os.path.join(UV_DIR, "sample_SampleA-1.Sample.asc"),
         os.path.join(UV_DIR, "sample_SampleA-2.Sample.asc")],
    ),
    (
        "SampleB",
        ["sample_SampleB-1", "sample_SampleB-2"],
        [os.path.join(UV_DIR, "sample_SampleB-1.Sample.asc"),
         os.path.join(UV_DIR, "sample_SampleB-2.Sample.asc")],
    ),
]

resultados_abs = {}

for nombre, hojas_ftir, archivos_uv in MUESTRAS_ABS:
    print(f"  Procesando {nombre}...")
    archivos_ir = {
        "path":          [FTIR_DIR],
        "path_muestras": {FTIR_DATA: hojas_ftir},
        "path_referencias": {FTIR_REF: ["zero_20260409", "base_20260409"]},
        "zero":    {FTIR_REF: "zero_20260409"},
        "base":    {FTIR_REF: "base_20260409"},
        "ventana": {},
    }
    archivos_uv_zero_base = {
        "ZeroLine": os.path.join(UV_DIR, "zero_demo.asc"),
        "BaseLine": os.path.join(UV_DIR, "base_demo.asc"),
        "ventana":  [],
    }
    m = Muestra(nombre, archivos_ir, archivos_uv_zero_base, archivos_uv,
                REFS_IR, REFS_UV, DATOS_BASICOS_ABS, "")

    df_ir      = m.procesar_datos_tfir()
    abs_ir     = m.medidas_ir(df_ir, False)
    df_uv      = m.leer_datos_UV(False)
    abs_uv, SWR, SWA, SWR_std = m.medidas_UV(df_uv, False)
    df_comb    = m.combinar_uv_ir(abs_ir, abs_uv)
    emitancia, df_abs = m.emitancia(df_comb)

    resultados_abs[nombre] = {
        "SWR": SWR, "SWA": SWA, "SWR_std": SWR_std,
        "emitancia": emitancia, "df": df_abs.copy(),
    }
    print(f"    SWA={SWA:.4f}  SWR={SWR:.4f}  SWR_std={SWR_std:.6f}  ε={emitancia:.4f}")


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 2 — REFLECTANCIA
#  Replica el cálculo que ejecuta la plantilla Excel del programa real.
#  Fórmula: R(λ) = (Iw - zero) / (base - zero) * r_ref
#  SWR = Σ R(λ)·ASTM(λ)   (misma ponderación que medidas_UV, muestra.py:305)
#  SWR_std: misma lógica que medidas_UV:287-291
# ═══════════════════════════════════════════════════════════════════════════════

print("\n── REFLECTANCIA ──")

paths_refl     = [os.path.join(REFL_DIR, "sample_ReflCSP-1.Sample.asc"),
                  os.path.join(REFL_DIR, "sample_ReflCSP-2.Sample.asc")]
path_zero_refl = os.path.join(REFL_DIR, "zero_refl.asc")
path_base_refl = os.path.join(REFL_DIR, "base_refl.asc")

df_refl = leer_asc(path_zero_refl, "zero").merge(
          leer_asc(path_base_refl, "base"), on="nm")
for i, p in enumerate(paths_refl, 1):
    df_refl = df_refl.merge(leer_asc(p, f"I{i}"), on="nm")

# Referencia desde hoja reflectores (escala %; se divide por 100 → fracción)
ref_refl = pd.read_excel(REFS_PATH, sheet_name="reflectores",
                         usecols=["wvl [nm]", "MASTER OMT"])
ref_refl.rename(columns={"wvl [nm]": "nm", "MASTER OMT": "r_ref"}, inplace=True)
ref_refl["r_ref"] /= 100.0
df_refl = df_refl.merge(ref_refl, on="nm", how="inner")

# Espectro solar ASTM G173 direct (normalizado, suma=1)
solar = leer_espectro_solar()
df_refl = df_refl.merge(
    solar[["nm", "ASTM_G173_direct"]].rename(columns={"ASTM_G173_direct": "ASTM"}),
    on="nm", how="inner")
df_refl.sort_values("nm", inplace=True, ignore_index=True)

intensity_cols_r = [f"I{i}" for i in range(1, len(paths_refl) + 1)]
df_refl["Iw"]   = df_refl[intensity_cols_r].mean(axis=1)
df_refl["refl"] = ((df_refl["Iw"] - df_refl["zero"]) /
                   (df_refl["base"] - df_refl["zero"])) * df_refl["r_ref"]
df_refl["refl"] = df_refl["refl"].clip(0, 1)

SWR_refl = float((df_refl["refl"] * df_refl["ASTM"]).sum())

refl_list_std = []
for col in intensity_cols_r:
    r_i = (((df_refl[col] - df_refl["zero"]) /
             (df_refl["base"] - df_refl["zero"])) * df_refl["r_ref"]) * df_refl["ASTM"].sum()
    refl_list_std.append(r_i)
SWR_std_refl = float(np.std(refl_list_std))

resultados_refl = {"SWR": SWR_refl, "SWR_std": SWR_std_refl, "df": df_refl.copy()}
print(f"  ReflCSP — SWR={SWR_refl:.4f}  SWR_std={SWR_std_refl:.6f}")


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 3 — TRANSMITANCIA CSP
#  Los archivos .asc contienen directamente la transmitancia (escala 0-1).
#  El programa real solo copia los datos al Excel; la plantilla calcula T_solar.
#  T_solar = Σ T(λ)·E891(λ)   (espectro directo, normalizado suma=1)
# ═══════════════════════════════════════════════════════════════════════════════

print("\n── TRANSMITANCIA CSP ──")


def calcular_transmitancia(paths_muestras, espectro_col):
    """
    Transmitancia solar ponderada.
    Replica la fórmula de la plantilla Excel del programa real.
    """
    dfs = [leer_asc(p, f"I{i}") for i, p in enumerate(paths_muestras, 1)]
    df = dfs[0].copy()
    for d in dfs[1:]:
        df = df.merge(d, on="nm", how="inner")

    solar = leer_espectro_solar()
    df = df.merge(solar[["nm", espectro_col]].rename(columns={espectro_col: "W"}),
                  on="nm", how="inner")
    df.sort_values("nm", inplace=True, ignore_index=True)

    icols = [c for c in df.columns if c.startswith("I")]
    df["T_mean"] = df[icols].mean(axis=1)
    T_solar = float((df["T_mean"] * df["W"]).sum())
    T_std   = float(np.std([(df[c] * df["W"]).sum() for c in icols]))
    return T_solar, T_std, df


paths_tcsp = [os.path.join(TCSP_DIR, "sample_GlassCSP-1.Sample.asc"),
              os.path.join(TCSP_DIR, "sample_GlassCSP-2.Sample.asc")]
T_csp, T_csp_std, df_tcsp = calcular_transmitancia(paths_tcsp, "E891")
resultados_tcsp = {"T_solar": T_csp, "T_std": T_csp_std, "espectro": "E891", "df": df_tcsp}
print(f"  GlassCSP — T_solar(E891)={T_csp:.4f}  T_std={T_csp_std:.6f}")


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 4 — TRANSMITANCIA PV
#  Igual que CSP pero ponderada con ASTM G173 global (espectro AM1.5G para PV).
# ═══════════════════════════════════════════════════════════════════════════════

print("\n── TRANSMITANCIA PV ──")

paths_tpv = [os.path.join(TPV_DIR, "sample_GlassPV-1.Sample.asc"),
             os.path.join(TPV_DIR, "sample_GlassPV-2.Sample.asc")]
T_pv, T_pv_std, df_tpv = calcular_transmitancia(paths_tpv, "ASTM_G173_global")
resultados_tpv = {"T_solar": T_pv, "T_std": T_pv_std, "espectro": "ASTM G173 global", "df": df_tpv}
print(f"  GlassPV — T_solar(ASTM global)={T_pv:.4f}  T_std={T_pv_std:.6f}")


# ═══════════════════════════════════════════════════════════════════════════════
#  HTML PLOTLY — 4 subgráficos
# ═══════════════════════════════════════════════════════════════════════════════

print("\n── Generando HTML ──")

COLORES = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728"]

fig = make_subplots(
    rows=2, cols=2,
    subplot_titles=[
        "Absortancia  α(λ)",
        "Reflectancia  ρ(λ)",
        "Transmitancia CSP  τ(λ)",
        "Transmitancia PV  τ(λ)",
    ],
    horizontal_spacing=0.10,
    vertical_spacing=0.16,
)

# ── Subplot 1: Absortancia — espectro completo UV+IR en µm ────────────────────
for i, (nombre, res) in enumerate(resultados_abs.items()):
    df = res["df"]
    fig.add_trace(go.Scatter(
        x=df["µm"], y=df["abs"].clip(0, 1),
        name=f"{nombre}  α  SWA={res['SWA']:.4f}  ε={res['emitancia']:.4f}",
        line=dict(color=COLORES[i], width=1.8),
        legendgroup="abs",
    ), row=1, col=1)
    fig.add_trace(go.Scatter(
        x=df["µm"], y=df["refl"].clip(0, 1),
        name=f"{nombre}  ρ",
        line=dict(color=COLORES[i], width=1, dash="dot"),
        legendgroup="abs", showlegend=True,
    ), row=1, col=1)

# Curva de Planck normalizada como referencia del rango térmico
h_p, c_p, k_p = 6.63e-34, 2.997e8, 1.38e-23
df0 = list(resultados_abs.values())[0]["df"]
lm  = df0["µm"].values * 1e-6
Mbb = (2 * np.pi * h_p * c_p**2) / (lm**5) / (np.exp(h_p * c_p / (lm * k_p * 373.0)) - 1)
fig.add_trace(go.Scatter(
    x=df0["µm"], y=Mbb / Mbb.max(),
    name="Planck 373 K (norm.)",
    line=dict(color="gray", width=1.2, dash="dash"),
    legendgroup="abs",
), row=1, col=1)

# ── Subplot 2: Reflectancia ───────────────────────────────────────────────────
df_r = resultados_refl["df"]
icols_r = [c for c in df_r.columns if c.startswith("I")]
for i, col in enumerate(icols_r):
    r_i = ((df_r[col] - df_r["zero"]) / (df_r["base"] - df_r["zero"])) * df_r["r_ref"]
    fig.add_trace(go.Scatter(
        x=df_r["nm"], y=r_i.clip(0, 1),
        name=f"ReflCSP réplica {i + 1}",
        line=dict(color=COLORES[i], width=1, dash="dot"),
        legendgroup="refl",
    ), row=1, col=2)
fig.add_trace(go.Scatter(
    x=df_r["nm"], y=df_r["refl"],
    name=f"ReflCSP media  SWR={resultados_refl['SWR']:.4f}",
    line=dict(color=COLORES[2], width=2),
    legendgroup="refl",
), row=1, col=2)

# ── Subplot 3: Transmitancia CSP ──────────────────────────────────────────────
df_c  = resultados_tcsp["df"]
icols_c = [c for c in df_c.columns if c.startswith("I")]
for i, col in enumerate(icols_c):
    fig.add_trace(go.Scatter(
        x=df_c["nm"], y=df_c[col].clip(0, 1),
        name=f"GlassCSP réplica {i + 1}",
        line=dict(color=COLORES[i], width=1, dash="dot"),
        legendgroup="tcsp",
    ), row=2, col=1)
fig.add_trace(go.Scatter(
    x=df_c["nm"], y=df_c["T_mean"].clip(0, 1),
    name=f"GlassCSP media  T(E891)={resultados_tcsp['T_solar']:.4f}",
    line=dict(color=COLORES[2], width=2),
    legendgroup="tcsp",
), row=2, col=1)

# ── Subplot 4: Transmitancia PV ───────────────────────────────────────────────
df_p  = resultados_tpv["df"]
icols_p = [c for c in df_p.columns if c.startswith("I")]
for i, col in enumerate(icols_p):
    fig.add_trace(go.Scatter(
        x=df_p["nm"], y=df_p[col].clip(0, 1),
        name=f"GlassPV réplica {i + 1}",
        line=dict(color=COLORES[i], width=1, dash="dot"),
        legendgroup="tpv",
    ), row=2, col=2)
fig.add_trace(go.Scatter(
    x=df_p["nm"], y=df_p["T_mean"].clip(0, 1),
    name=f"GlassPV media  T(ASTM glob.)={resultados_tpv['T_solar']:.4f}",
    line=dict(color=COLORES[2], width=2),
    legendgroup="tpv",
), row=2, col=2)

# ── Ejes ─────────────────────────────────────────────────────────────────────
fig.update_xaxes(title_text="λ (µm)", row=1, col=1)
fig.update_xaxes(title_text="λ (nm)", row=1, col=2)
fig.update_xaxes(title_text="λ (nm)", row=2, col=1)
fig.update_xaxes(title_text="λ (nm)", row=2, col=2)
fig.update_yaxes(title_text="α , ρ", range=[0, 1.05], row=1, col=1)
fig.update_yaxes(title_text="ρ",     range=[0, 1.05], row=1, col=2)
fig.update_yaxes(title_text="τ",     range=[0, 1.05], row=2, col=1)
fig.update_yaxes(title_text="τ",     range=[0, 1.05], row=2, col=2)

fig.update_layout(
    title_text="Aliquindoi — Demo de resultados espectrales",
    title_font_size=15,
    height=780,
    legend=dict(orientation="v", x=1.01, y=1, font_size=10),
    font=dict(size=11),
)

fig.write_html(HTML_OUT)
print(f"  Guardado: {HTML_OUT}")


# ═══════════════════════════════════════════════════════════════════════════════
#  RESUMEN FINAL
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("RESUMEN FINAL DE RESULTADOS")
print("=" * 60)

print(f"\n{'Muestra':<12} {'Tipo':<16} {'SWA':>7} {'SWR':>7} {'SWR_std':>10} {'ε(373K)':>8}")
print("-" * 64)
for nombre, res in resultados_abs.items():
    print(f"{nombre:<12} {'Absortancia':<16} "
          f"{res['SWA']:>7.4f} {res['SWR']:>7.4f} "
          f"{res['SWR_std']:>10.6f} {res['emitancia']:>8.4f}")

print(f"\n{'Muestra':<12} {'Tipo':<16} {'SWR':>7} {'SWR_std':>10} {'Referencia'}")
print("-" * 55)
print(f"{'ReflCSP':<12} {'Reflectancia':<16} "
      f"{resultados_refl['SWR']:>7.4f} {resultados_refl['SWR_std']:>10.6f}  MASTER OMT")

print(f"\n{'Muestra':<12} {'Tipo':<16} {'T_solar':>8} {'T_std':>10}  Espectro")
print("-" * 60)
print(f"{'GlassCSP':<12} {'Trans. CSP':<16} "
      f"{resultados_tcsp['T_solar']:>8.4f} {resultados_tcsp['T_std']:>10.6f}  E891")
print(f"{'GlassPV':<12} {'Trans. PV':<16} "
      f"{resultados_tpv['T_solar']:>8.4f} {resultados_tpv['T_std']:>10.6f}  ASTM G173 global")

print("\n" + "=" * 60)
