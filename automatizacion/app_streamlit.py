#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FASE 4 - Interfaz web (Streamlit)
=================================

Dos modos de trabajo:

1. Carpeta local: el usuario indica una ruta (p. ej. Descargas\\Nueva carpeta (2))
   y la app ejecuta el pipeline directamente sobre esos archivos, sin subirlos.
2. Subida: cinco file_uploader; los archivos se guardan en `subidas/`.

En ambos casos se ejecutan las Fases 1, 2 y 5 y se ofrecen tres descargas.
La salida del proceso se lee en tiempo real (Popen + cola) para evitar
bloqueos de buffer en Windows.

Arranque:
    PYTHONPATH=/tmp/my_py python3 -m streamlit run automatizacion/app_streamlit.py
"""

import os
import queue
import shutil
import subprocess
import sys
import threading
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

NOMBRES_LIMPIOS = {
    "Seguimiento E2E VODAFONE.xlsx": "seguimiento.xlsx",
    "JIRA con DR.xlsx": "jira.xlsx",
    "Hoja de cálculo sin título.xlsx": "hoja.xlsx",
    "MinisterioPowerBi.csv": "ministerio.csv",
    "plantilla_previsiones.xlsx": "plantilla.xlsx",
}

ARCHIVOS_CARPETA = [
    "Seguimiento E2E VODAFONE.xlsx",
    "JIRA con DR.xlsx",
    "Hoja de cálculo sin título.xlsx",
    "MinisterioPowerBi.csv",
    "plantilla_previsiones.xlsx",
]
REVISION = "Revision Legalizacion SEMI SemN.xlsx"

st.set_page_config(page_title="Automatización de Informes SEMI",
                   page_icon="📊", layout="wide")

st.title("📊 Automatización de Informes SEMI")
st.write(
    "Puedes indicar una **carpeta local** con los archivos (recomendado en "
    "Windows, evita subir ~100 MB) o subirlos aquí. Se ejecutarán las "
    "**Fases 1, 2 y 5** y se generarán los tres informes."
)


# ----------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------
def entorno_hijos() -> dict:
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    env["PYTHONUNBUFFERED"] = "1"
    vendor = Path("/tmp/my_py")
    if vendor.is_dir():
        actual = env.get("PYTHONPATH", "")
        if str(vendor) not in actual.split(os.pathsep):
            env["PYTHONPATH"] = str(vendor) + (os.pathsep + actual if actual else "")
    return env


def limpiar_subidas() -> None:
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


def _avance(linea: str, actual: int) -> int:
    if "FASE 1" in linea:
        return 10
    if "FASE 2" in linea:
        return 40
    if "FASE 5" in linea:
        return 65
    if "[OK] Proceso completado" in linea:
        return 100
    return actual


def ejecutar_con_progreso(cmd: list, timeout: int = 3600) -> str:
    """Lanza main.py y transmite su salida en vivo (sin bloquear el pipe)."""
    proceso = subprocess.Popen(
        cmd, cwd=str(BASE), env=entorno_hijos(),
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, bufsize=1,
    )
    cola: queue.Queue = queue.Queue()

    def lector():
        try:
            for linea in iter(proceso.stdout.readline, ""):
                cola.put(linea)
        finally:
            cola.put(None)

    threading.Thread(target=lector, daemon=True).start()

    lineas = []
    inicio = time.time()
    ultimo = 0.0
    avance = 0
    try:
        with st.status("Procesando datos...", expanded=True) as estado:
            barra = st.progress(0, text="Iniciando el pipeline...")
            while True:
                try:
                    linea = cola.get(timeout=0.5)
                except queue.Empty:
                    if time.time() - inicio > timeout:
                        proceso.kill()
                        raise TimeoutError(
                            f"El proceso superó {timeout // 60} minutos."
                        )
                    continue

                if linea is None:
                    break

                linea = linea.rstrip()
                if not linea:
                    continue
                lineas.append(linea)
                avance = _avance(linea, avance)

                if time.time() - ultimo > 0.4:
                    barra.progress(min(avance, 99), text=linea[:120])
                    estado.update(label=f"Procesando... {linea[:90]}")
                    ultimo = time.time()

            codigo = proceso.wait()
            if codigo != 0:
                raise RuntimeError(f"El pipeline terminó con código {codigo}.")
            barra.progress(100, text="Proceso completado")
            estado.update(label="Proceso completado", state="complete")
    except Exception:
        if proceso.poll() is None:
            proceso.kill()
        raise
    finally:
        try:
            proceso.stdout.close()
        except Exception:
            pass
    return "\n".join(lineas)


def leer_salidas(carpeta: Path) -> dict:
    archivos = {
        "seg": carpeta / "seguimiento.xlsx",
        "legal": carpeta / "Legal_Final.xlsx",
        "prev": carpeta / "Seguimiento_Previsiones_Limpio.xlsx",
    }
    faltan = [p.name for p in archivos.values() if not p.is_file()]
    if faltan:
        raise RuntimeError("No se generaron: " + ", ".join(faltan))
    return {k: p.read_bytes() for k, p in archivos.items()}


# ----------------------------------------------------------------------
# 1 · Origen de los datos
# ----------------------------------------------------------------------
st.subheader("1 · Origen de los datos")
ruta_texto = st.text_input(
    "Ruta local de la carpeta de trabajo (opcional)",
    placeholder=r"C:\Users\describa\Downloads\Nueva carpeta (2)",
    help="Si la indicas, la app usa esos archivos y no sube nada. "
         "Debe contener los 5 archivos y, preferiblemente, la Revision.",
)

ruta = Path(ruta_texto.strip()).expanduser() if ruta_texto.strip() else None
modo_ruta = ruta is not None
faltan_carpeta: list = []
revision_local = None

if modo_ruta:
    if not ruta.is_dir():
        st.error(f"La carpeta no existe: `{ruta}`")
        faltan_carpeta = ["(carpeta no encontrada)"]
    else:
        faltan_carpeta = [n for n in ARCHIVOS_CARPETA if not (ruta / n).is_file()]
        rev = ruta / REVISION
        revision_local = rev if rev.is_file() else None
        if faltan_carpeta:
            st.error("Faltan en la carpeta: " + ", ".join(faltan_carpeta))
        else:
            st.success(f"5 archivos encontrados en `{ruta}`")
        if revision_local is None:
            st.warning(
                f"No se encontró `{REVISION}`; se usará la de la carpeta del proyecto."
            )
else:
    st.caption("O sube los archivos aquí:")
    col1, col2 = st.columns(2)
    with col1:
        up_seg = st.file_uploader("Seguimiento E2E VODAFONE.xlsx", type=["xlsx"], key="seg")
        up_jira = st.file_uploader("JIRA con DR.xlsx", type=["xlsx"], key="jira")
        up_plantilla = st.file_uploader("plantilla_previsiones.xlsx", type=["xlsx"],
                                        key="plantilla")
    with col2:
        up_hoja = st.file_uploader("Hoja de cálculo sin título.xlsx", type=["xlsx"],
                                   key="hoja")
        up_csv = st.file_uploader("MinisterioPowerBi.csv", type=["csv"], key="csv")

    subidos = {
        "Seguimiento E2E VODAFONE.xlsx": up_seg,
        "JIRA con DR.xlsx": up_jira,
        "Hoja de cálculo sin título.xlsx": up_hoja,
        "MinisterioPowerBi.csv": up_csv,
        "plantilla_previsiones.xlsx": up_plantilla,
    }
    faltan_subida = [n for n, f in subidos.items() if f is None]
    if faltan_subida:
        st.info("Faltan por subir: " + ", ".join(faltan_subida))
    else:
        st.success("Los 5 archivos están listos.")


# ----------------------------------------------------------------------
# 2 · Ejecución
# ----------------------------------------------------------------------
if "resultado" not in st.session_state:
    st.session_state.resultado = None

st.subheader("2 · Ejecución")
pulsado = st.button("🚀 Ejecutar Proceso", type="primary", width="stretch")

if pulsado:
    st.session_state.resultado = None
    error = None
    carpeta_salida = None
    limpiar = False

    try:
        if modo_ruta:
            if faltan_carpeta:
                raise RuntimeError("Revisa la carpeta: " + ", ".join(faltan_carpeta))
            carpeta_salida = ruta / "resultados_app"
            carpeta_salida.mkdir(parents=True, exist_ok=True)
            cmd = [
                sys.executable, str(MAIN),
                f"--seguimiento={ruta / 'Seguimiento E2E VODAFONE.xlsx'}",
                f"--jira={ruta / 'JIRA con DR.xlsx'}",
                f"--hoja={ruta / 'Hoja de cálculo sin título.xlsx'}",
                f"--csv={ruta / 'MinisterioPowerBi.csv'}",
                f"--plantilla={ruta / 'plantilla_previsiones.xlsx'}",
                f"--jira-out={carpeta_salida / 'JIRA.xlsx'}",
                f"--out={carpeta_salida}",
            ]
            if revision_local is not None:
                cmd.append(f"--revision={revision_local}")
        else:
            if faltan_subida:
                raise RuntimeError("Faltan archivos por subir: "
                                   + ", ".join(faltan_subida))
            SUBIDAS.mkdir(parents=True, exist_ok=True)
            SALIDA.mkdir(parents=True, exist_ok=True)
            limpiar = True
            rutas = {}
            for etiqueta, archivo in subidos.items():
                destino = SUBIDAS / NOMBRES_LIMPIOS[etiqueta]
                destino.write_bytes(archivo.getvalue())
                rutas[etiqueta] = destino
            carpeta_salida = SALIDA
            cmd = [
                sys.executable, str(MAIN),
                f"--seguimiento={rutas['Seguimiento E2E VODAFONE.xlsx']}",
                f"--jira={rutas['JIRA con DR.xlsx']}",
                f"--hoja={rutas['Hoja de cálculo sin título.xlsx']}",
                f"--csv={rutas['MinisterioPowerBi.csv']}",
                f"--plantilla={rutas['plantilla_previsiones.xlsx']}",
                f"--jira-out={SALIDA / 'JIRA.xlsx'}",
                f"--out={SALIDA}",
            ]

        t0 = time.time()
        log = ejecutar_con_progreso(cmd, timeout=3600)
        datos = leer_salidas(carpeta_salida)
        datos.update({"elapsed": time.time() - t0, "log": log})
        st.session_state.resultado = datos

    except TimeoutError as e:
        error = str(e)
    except Exception as e:  # noqa: BLE001
        error = str(e)
    finally:
        if limpiar:
            limpiar_subidas()

    if error:
        st.error(error)
        st.caption("Revisa el registro o ejecuta `main.py` en terminal para ver el detalle.")


# ----------------------------------------------------------------------
# 3 · Resultados
# ----------------------------------------------------------------------
res = st.session_state.resultado
if res:
    st.subheader("3 · Resultados")
    st.success(f"✅ Proceso completado en {res['elapsed']:.1f} s")

    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            "⬇️ seguimiento.xlsx",
            data=res["seg"],
            file_name="seguimiento.xlsx",
            mime=MIME_XLSX,
            width="stretch",
            key="dl_seg",
        )
        st.caption(f"{len(res['seg']) / 1e6:.2f} MB · NUEVO SEGUIMIENTO")
    with d2:
        st.download_button(
            "⬇️ Legal_Final.xlsx",
            data=res["legal"],
            file_name="Legal_Final.xlsx",
            mime=MIME_XLSX,
            width="stretch",
            key="dl_legal",
        )
        st.caption(f"{len(res['legal']) / 1e6:.2f} MB · LEGAL_VDF / LEGAL_OSP")
    with d3:
        st.download_button(
            "⬇️ Seguimiento_Previsiones_Limpio.xlsx",
            data=res["prev"],
            file_name="Seguimiento_Previsiones_Limpio.xlsx",
            mime=MIME_XLSX,
            width="stretch",
            key="dl_prev",
        )
        st.caption(f"{len(res['prev']) / 1e6:.2f} MB · 7 hojas")

    with st.expander("Ver registro de ejecución"):
        st.code(res["log"], language="text")

st.divider()
st.caption(
    "Fase 1: actualiza el Seguimiento E2E. Fase 2: genera el Legal VDF/OSP. "
    "Fase 5: construye el Seguimiento de Previsiones limpio (ZOOM incluido)."
)