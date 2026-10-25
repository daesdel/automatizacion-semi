#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FASE 4 - Interfaz web (Streamlit)
=================================

Interfaz para lanzar la actualización completa sin usar la terminal:
  - Sube los 4 archivos de entrada.
  - Los guarda en `subidas/` (carpeta local del proyecto, no en %TEMP%).
  - Ejecuta `main.py` pasándole esas rutas.
  - Permite descargar seguimiento.xlsx y Legal_Final.xlsx.
  - Borra los archivos de `subidas/` al terminar (éxito o error).

Arranque:
    cd /Users/user123/Library/CloudStorage/OneDrive-Personal/Documentos/semi
    PYTHONPATH=/tmp/my_py python3 -m streamlit run automatizacion/app_streamlit.py
"""

import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import streamlit as st

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
MAIN = HERE / "main.py"
SUBIDAS = BASE / "subidas"
SALIDA = SUBIDAS / "resultados_app"

MIME_XLSX = ("application/vnd.openxmlformats-officedocument."
             "spreadsheetml.sheet")

# Etiqueta visible -> nombre limpio con el que se guarda en subidas/
NOMBRES_LIMPIOS = {
    "Seguimiento E2E VODAFONE.xlsx": "seguimiento.xlsx",
    "JIRA con DR.xlsx": "jira.xlsx",
    "Hoja de cálculo sin título.xlsx": "hoja.xlsx",
    "MinisterioPowerBi.csv": "ministerio.csv",
    "plantilla_previsiones.xlsx": "plantilla.xlsx",
}

# ----------------------------------------------------------------------
# Página
# ----------------------------------------------------------------------
st.set_page_config(page_title="Automatización de Informes SEMI",
                   page_icon="📊", layout="wide")

st.title("📊 Automatización de Informes SEMI")
st.write(
    "Sube los cinco archivos de entrada y pulsa **🚀 Ejecutar Proceso**. "
    "Se ejecutarán las Fases 1, 2 y 5 y se generarán los tres informes. "
    "Los archivos se guardan temporalmente en `subidas/` y se eliminan al terminar."
)


# ----------------------------------------------------------------------
# Entorno de los subprocesos
# ----------------------------------------------------------------------
def entorno_hijos() -> dict:
    env = os.environ.copy()
    # Texto en UTF-8 también en la consola de Windows
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    # Permite lanzar sin PYTHONPATH en macOS si las librerías están en /tmp/my_py
    vendor = Path("/tmp/my_py")
    if vendor.is_dir():
        actual = env.get("PYTHONPATH", "")
        if str(vendor) not in actual.split(os.pathsep):
            env["PYTHONPATH"] = str(vendor) + (os.pathsep + actual if actual else "")
    return env


def limpiar_subidas() -> None:
    """Borra el contenido de subidas/ (se conserva la carpeta)."""
    if not SUBIDAS.is_dir():
        return
    for p in SUBIDAS.iterdir():
        try:
            if p.is_dir():
                shutil.rmtree(p, ignore_errors=True)
            else:
                p.unlink()
        except OSError:
            pass


# ----------------------------------------------------------------------
# Carga de archivos
# ----------------------------------------------------------------------
st.subheader("1 · Archivos de entrada")
col1, col2 = st.columns(2)

with col1:
    up_seg = st.file_uploader("Seguimiento E2E VODAFONE.xlsx", type=["xlsx"],
                              key="seg", help="Libro de seguimiento (21 MB aprox.)")
    up_jira = st.file_uploader("JIRA con DR.xlsx", type=["xlsx"],
                               key="jira", help="Export de JIRA (35 MB aprox.)")
with col2:
    up_hoja = st.file_uploader("Hoja de cálculo sin título.xlsx", type=["xlsx"],
                               key="hoja", help="Consolidado de JIRA (34 MB aprox.)")
    up_csv = st.file_uploader("MinisterioPowerBi.csv", type=["csv"],
                              key="csv", help="CSV del Ministerio (23 MB aprox.)")
    up_plantilla = st.file_uploader("plantilla_previsiones.xlsx", type=["xlsx"],
                                    key="plantilla",
                                    help="Plantilla de PREVISIONES JUAN (Fase 5)")

subidos = {
    "Seguimiento E2E VODAFONE.xlsx": up_seg,
    "JIRA con DR.xlsx": up_jira,
    "Hoja de cálculo sin título.xlsx": up_hoja,
    "MinisterioPowerBi.csv": up_csv,
    "plantilla_previsiones.xlsx": up_plantilla,
}
faltan = [n for n, f in subidos.items() if f is None]

if faltan:
    st.info("Faltan por subir: " + ", ".join(faltan))
else:
    st.success("Los 5 archivos están listos. Puedes ejecutar el proceso.")

# ----------------------------------------------------------------------
# Ejecución
# ----------------------------------------------------------------------
if "resultado" not in st.session_state:
    st.session_state.resultado = None

st.subheader("2 · Ejecución")
pulsado = st.button("🚀 Ejecutar Proceso", type="primary", use_container_width=True)


def cola(texto: str, lineas: int = 30) -> str:
    return "\n".join((texto or "").strip().splitlines()[-lineas:])


if pulsado:
    st.session_state.resultado = None

    if faltan:
        st.error("No se puede ejecutar: falta por subir " + ", ".join(faltan) + ".")
    else:
        SUBIDAS.mkdir(parents=True, exist_ok=True)
        SALIDA.mkdir(parents=True, exist_ok=True)
        proceso = None
        try:
            rutas = {}
            for etiqueta, archivo in subidos.items():
                destino = SUBIDAS / NOMBRES_LIMPIOS[etiqueta]
                destino.write_bytes(archivo.getvalue())
                rutas[etiqueta] = destino

            t0 = time.time()
            cmd = [
                sys.executable, str(MAIN),
                f"--seguimiento={rutas['Seguimiento E2E VODAFONE.xlsx']}",
                f"--jira={rutas['JIRA con DR.xlsx']}",
                f"--hoja={rutas['Hoja de cálculo sin título.xlsx']}",
                f"--csv={rutas['MinisterioPowerBi.csv']}",
                f"--plantilla={rutas['plantilla_previsiones.xlsx']}",
                f"--out={SALIDA}",
            ]

            with st.spinner("Procesando datos... (puede tardar varios minutos)"):
                proceso = subprocess.run(
                    cmd, cwd=str(BASE), env=entorno_hijos(),
                    capture_output=True, text=True, timeout=1800,
                )
                if proceso.returncode != 0:
                    raise RuntimeError(
                        "El proceso terminó con errores (código "
                        f"{proceso.returncode})."
                    )

                f_seg = SALIDA / "seguimiento.xlsx"
                f_legal = SALIDA / "Legal_Final.xlsx"
                f_prev = SALIDA / "Seguimiento_Previsiones_Limpio.xlsx"
                if not (f_seg.is_file() and f_legal.is_file() and f_prev.is_file()):
                    raise RuntimeError(
                        "El proceso terminó pero no se generaron los tres archivos esperados."
                    )

                st.session_state.resultado = {
                    "seg": f_seg.read_bytes(),
                    "legal": f_legal.read_bytes(),
                    "prev": f_prev.read_bytes(),
                    "elapsed": time.time() - t0,
                    "log": (proceso.stdout or "") + "\n" + (proceso.stderr or ""),
                }

        except subprocess.TimeoutExpired:
            st.error("❌ El proceso superó el tiempo máximo permitido (30 minutos).")
        except Exception as e:  # noqa: BLE001
            st.error(f"❌ {e}")
            if proceso is not None:
                st.code(cola((proceso.stdout or "") + "\n" + (proceso.stderr or "")),
                        language="text")
        finally:
            limpiar_subidas()

# ----------------------------------------------------------------------
# Resultados y descargas
# ----------------------------------------------------------------------
res = st.session_state.resultado
if res:
    st.subheader("3 · Resultados")
    st.success(f"✅ Proceso completado en {res['elapsed']:.1f} s")

    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            "⬇️ Descargar seguimiento.xlsx",
            data=res["seg"],
            file_name="seguimiento.xlsx",
            mime=MIME_XLSX,
            use_container_width=True,
            key="dl_seg",
        )
        st.caption(f"{len(res['seg']) / 1e6:.2f} MB · hoja NUEVO SEGUIMIENTO")
    with d2:
        st.download_button(
            "⬇️ Descargar Legal_Final.xlsx",
            data=res["legal"],
            file_name="Legal_Final.xlsx",
            mime=MIME_XLSX,
            use_container_width=True,
            key="dl_legal",
        )
        st.caption(f"{len(res['legal']) / 1e6:.2f} MB · hojas LEGAL_VDF y LEGAL_OSP")
    with d3:
        st.download_button(
            "⬇️ Descargar Seguimiento_Previsiones_Limpio.xlsx",
            data=res["prev"],
            file_name="Seguimiento_Previsiones_Limpio.xlsx",
            mime=MIME_XLSX,
            use_container_width=True,
            key="dl_prev",
        )
        st.caption(f"{len(res['prev']) / 1e6:.2f} MB · 7 hojas")

    with st.expander("Ver registro de ejecución"):
        st.code(res["log"], language="text")

st.divider()
st.caption(
    "Fase 1: actualiza VOLCADO JIRA desde JIRA con DR + Hoja sin título y "
    "recalcula NUEVO SEGUIMIENTO. Fase 2: procesa el CSV del Ministerio y "
    "genera la matriz VDF/OSP."
)