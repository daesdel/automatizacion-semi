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
from datetime import datetime
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
    "Seguimiento_Previsiones_Limpio_anterior.xlsx": "anterior.xlsx",
}

ARCHIVOS_CARPETA = [
    "Seguimiento E2E VODAFONE.xlsx",
    "JIRA con DR.xlsx",
    "Hoja de cálculo sin título.xlsx",
    "MinisterioPowerBi.csv",
]
REVISION = "Revision Legalizacion SEMI SemN.xlsx"

st.set_page_config(page_title="Automatización de Informes SEMI",
                   page_icon="📊", layout="wide")

st.title("📊 Automatización de Informes SEMI")
st.write(
    "Puedes indicar una **carpeta local** con los archivos (recomendado en "
    "Windows, evita subir ~100 MB) o subirlos aquí. Se ejecutarán las "
    "**Fases 1, 2 y 5**. Opcionalmente, aporta el "
    "`Seguimiento_Previsiones_Limpio.xlsx` de la semana anterior para heredar "
    "las fórmulas de PREVISIONES JUAN."
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


def ejecutar_con_progreso(cmd: list, timeout: int = 3600) -> str:
    """Lanza main.py y muestra progreso, cronómetro y consola en vivo."""
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

    inicio_dt = datetime.now()
    inicio = time.time()
    lineas = []
    avance = 0.0
    tope = 5.0
    ultimo = 0.0

    def refrescar(barra, reloj, consola, texto):
        trans = time.time() - inicio
        reloj.markdown(
            f"**Inicio:** {inicio_dt:%d/%m/%Y %H:%M:%S}  ·  "
            f"**Transcurrido:** {int(trans // 60):02d}:{int(trans % 60):02d}"
        )
        if lineas:
            consola.code("\n".join(lineas[-14:]), language="text")
        barra.progress(min(int(avance), 99), text=texto[:120])

    try:
        with st.status("Procesando datos...", expanded=True) as estado:
            barra = st.progress(0, text="Iniciando el pipeline...")
            reloj = st.empty()
            consola = st.empty()
            refrescar(barra, reloj, consola, "Iniciando...")

            while True:
                try:
                    linea = cola.get(timeout=0.5)
                except queue.Empty:
                    if time.time() - inicio > timeout:
                        proceso.kill()
                        raise TimeoutError(
                            f"El proceso superó {timeout // 60} minutos."
                        )
                    if time.time() - ultimo > 0.5:
                        refrescar(barra, reloj, consola,
                                  lineas[-1] if lineas else "Esperando salida...")
                        ultimo = time.time()
                    continue

                if linea is None:
                    break

                linea = linea.rstrip()
                if not linea:
                    continue
                lineas.append(linea)

                if "FASE 1" in linea:
                    avance, tope = 5, 35
                elif "FASE 2" in linea:
                    avance, tope = max(avance, 35), 60
                elif "FASE 5" in linea:
                    avance, tope = max(avance, 60), 95
                elif "[OK] Proceso completado" in linea:
                    avance, tope = 100, 100
                else:
                    avance = min(tope, avance + 0.2)

                if time.time() - ultimo > 0.35:
                    refrescar(barra, reloj, consola, linea)
                    estado.update(label=f"Procesando... {linea[:90]}")
                    ultimo = time.time()

            codigo = proceso.wait()
            if codigo != 0:
                raise RuntimeError(f"El pipeline terminó con código {codigo}.")
            avance = 100
            refrescar(barra, reloj, consola, "Proceso completado")
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
    legales = sorted(carpeta.glob("*_legal ministerio.xlsx"),
                     key=lambda p: p.stat().st_mtime)
    archivos = {
        "seg": carpeta / "seguimiento.xlsx",
        "prev": carpeta / "Seguimiento_Previsiones_Limpio.xlsx",
    }
    faltan = [p.name for p in archivos.values() if not p.is_file()]
    if not legales:
        faltan.append("*_legal ministerio.xlsx")
    if faltan:
        raise RuntimeError("No se generaron: " + ", ".join(faltan))
    datos = {k: p.read_bytes() for k, p in archivos.items()}
    datos["legal"] = legales[-1].read_bytes()
    datos["legal_name"] = legales[-1].name
    return datos


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
anterior_txt = st.text_input(
    "Archivo anterior Seguimiento_Previsiones_Limpio.xlsx (opcional)",
    placeholder=r"C:\...\Seguimiento_Previsiones_Limpio.xlsx",
    help="Si no se indica, se generará desde cero SIN heredar fórmulas.",
)
anterior_path = Path(anterior_txt.strip()).expanduser() if anterior_txt.strip() else None
revision_local = None
anterior_local = None
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
        if anterior_path and anterior_path.is_file():
            anterior_local = anterior_path
        elif (ruta / "Seguimiento_Previsiones_Limpio.xlsx").is_file():
            anterior_local = ruta / "Seguimiento_Previsiones_Limpio.xlsx"
        if anterior_local is None:
            st.warning("Sin archivo anterior: PREVISIONES se generará desde cero, "
                       "sin heredar fórmulas.")
        if faltan_carpeta:
            st.error("Faltan en la carpeta: " + ", ".join(faltan_carpeta))
        else:
            st.success(f"Archivos base encontrados en `{ruta}`")
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
        up_anterior = st.file_uploader(
            "Seguimiento_Previsiones_Limpio.xlsx (semana anterior, opcional)",
            type=["xlsx"], key="anterior")
    with col2:
        up_hoja = st.file_uploader("Hoja de cálculo sin título.xlsx", type=["xlsx"],
                                   key="hoja")
        up_csv = st.file_uploader("MinisterioPowerBi.csv", type=["csv"], key="csv")

    subidos = {
        "Seguimiento E2E VODAFONE.xlsx": up_seg,
        "JIRA con DR.xlsx": up_jira,
        "Hoja de cálculo sin título.xlsx": up_hoja,
        "MinisterioPowerBi.csv": up_csv,
        "Seguimiento_Previsiones_Limpio_anterior.xlsx": up_anterior,
    }
    faltan_subida = [n for n, f in subidos.items()
                     if f is None and "anterior" not in n]
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
                f"--jira-out={carpeta_salida / 'JIRA.xlsx'}",
                f"--out={carpeta_salida}",
            ]
            jira_path = ruta / "JIRA con DR.xlsx"
            hoja_path = ruta / "Hoja de cálculo sin título.xlsx"
            if anterior_local is not None:
                cmd.append(f"--anterior={anterior_local}")
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
                if archivo is None:
                    continue
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
                f"--jira-out={SALIDA / 'JIRA.xlsx'}",
                f"--out={SALIDA}",
            ]
            jira_path = rutas["JIRA con DR.xlsx"]
            hoja_path = rutas["Hoja de cálculo sin título.xlsx"]
            if up_anterior is not None:
                cmd.append(f"--anterior={rutas['Seguimiento_Previsiones_Limpio_anterior.xlsx']}")

        t0 = time.time()
        log = ejecutar_con_progreso(cmd, timeout=3600)
        datos = leer_salidas(carpeta_salida)
        datos.update({"elapsed": time.time() - t0, "log": log,
                      "carpeta": str(carpeta_salida),
                      "jira": str(jira_path), "hoja": str(hoja_path)})
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
            f"⬇️ {res.get('legal_name', 'legal ministerio.xlsx')}",
            data=res["legal"],
            file_name=res.get("legal_name", "legal ministerio.xlsx"),
            mime=MIME_XLSX,
            width="stretch",
            key="dl_legal",
        )
        st.caption(f"{len(res['legal']) / 1e6:.2f} MB · legal_vdf / legal_osp")
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

# ----------------------------------------------------------------------
# 4 · Auditoría de tickets
# ----------------------------------------------------------------------
st.subheader("4 · Auditoría de tickets")
if st.button("🔎 Ejecutar Auditoría de Tickets", width="stretch"):
    base_audit = Path(res["carpeta"]) if res else None
    if base_audit is None or not (base_audit / "Seguimiento_Previsiones_Limpio.xlsx").is_file():
        st.warning("Ejecuta primero el pipeline para tener un PREVISIONES actual.")
    else:
        jira_a = Path(res["jira"])
        hoja_a = Path(res["hoja"])
        if not jira_a.is_file():
            jira_a = BASE / "Nueva carpeta (2)" / "JIRA con DR.xlsx"
        if not hoja_a.is_file():
            hoja_a = BASE / "Nueva carpeta (2)" / "Hoja de cálculo sin título.xlsx"
        salida_audit = base_audit / "Auditoria_Tickets_Faltantes.xlsx"
        cmd = [
            sys.executable, str(HERE / "fase6_auditoria.py"),
            f"--jira={jira_a}", f"--hoja={hoja_a}",
            f"--seguimiento={base_audit / 'seguimiento.xlsx'}",
            f"--prev={base_audit / 'Seguimiento_Previsiones_Limpio.xlsx'}",
            f"--out={salida_audit}",
        ]
        with st.spinner("Buscando IRs huérfanos..."):
            p = subprocess.run(cmd, cwd=str(BASE), env=entorno_hijos(),
                               capture_output=True, text=True, timeout=900)
        if p.returncode != 0:
            st.error("La auditoría falló.")
            st.code((p.stdout or "") + "\n" + (p.stderr or ""), language="text")
        elif salida_audit.is_file():
            st.success("Auditoría completada.")
            st.download_button(
                "⬇️ Auditoria_Tickets_Faltantes.xlsx",
                data=salida_audit.read_bytes(),
                file_name="Auditoria_Tickets_Faltantes.xlsx",
                mime=MIME_XLSX, width="stretch", key="dl_audit",
            )
