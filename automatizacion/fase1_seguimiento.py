#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FASE 1 - Seguimiento E2E VODAFONE  (pipeline optimizado)
========================================================

Optimizaciones aplicadas:
  1. Lectura con calamine (motor rápido, sin openpyxl salvo fallback).
  2. Escritura con xlsxwriter en modo constant_memory (write_only):
     solo la hoja NUEVO SEGUIMIENTO, que es el entregable.
  3. Sin merge de 261 columnas: el refresco de VOLCADO JIRA (opcional,
     --refresh) usa únicamente las 22 columnas que consumen los VLOOKUP.

Uso:
  python fase1_seguimiento.py              # recalcula sobre el VOLCADO JIRA actual
  python fase1_seguimiento.py --refresh    # además refresca VOLCADO JIRA desde
                                           # JIRA con DR + Hoja sin título (IR_VDF)

Salida:
  Nueva carpeta (2)/resultados/seguimiento.xlsx  (hoja NUEVO SEGUIMIENTO, valores)
"""

import gc
import re
import sys
import time
import warnings
from datetime import date, datetime
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
import xlsxwriter
from openpyxl.utils import column_index_from_string

warnings.filterwarnings("ignore", category=UserWarning)

# Consola de Windows (cp1252): evita UnicodeEncodeError al imprimir símbolos.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ----------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------
def _arg_path(nombre: str, defecto: Path) -> Path:
    """Ruta opcional por argumento: --nombre=/ruta/al/archivo."""
    pref = f"--{nombre}="
    for a in sys.argv[1:]:
        if a.startswith(pref):
            return Path(a.split("=", 1)[1]).expanduser()
    return defecto


BASE_DIR = Path(__file__).parent.parent / "Nueva carpeta (2)"
FICHERO = _arg_path("seguimiento", BASE_DIR / "Seguimiento E2E VODAFONE.xlsx")
JIRA_DR = _arg_path("jira", BASE_DIR / "JIRA con DR.xlsx")
HOJA_SIN_TITULO = _arg_path("hoja", BASE_DIR / "Hoja de cálculo sin título.xlsx")
OUT_DIR = _arg_path("out", BASE_DIR / "resultados")
OUT_DIR.mkdir(parents=True, exist_ok=True)

HOJA_NS = "NUEVO SEGUIMIENTO"
HOJA_VJ = "VOLCADO JIRA"
HOJA_EZ = "Encendidos ZR"

REFRESCAR = "--refresh" in sys.argv
FILAS = next((int(a.split("=")[1]) for a in sys.argv if a.startswith("--filas=")), None)

try:
    import python_calamine  # noqa: F401
    ENGINE = "calamine"
except ImportError:
    ENGINE = None

if ENGINE is None:
    print("[AVISO] python-calamine no esta instalado: se usara openpyxl "
          "(mucho mas lento, puede tardar varios minutos).")
    print("        Windows:  py -m pip install python-calamine")


try:
    import psutil
    _PROC = psutil.Process()
except Exception:          # psutil no instalado (o sin permisos)
    _PROC = None


def rss_mb() -> float:
    """Memoria residente del proceso en MB, multiplataforma.

    - psutil (Windows/macOS/Linux) si está instalado;
    - `resource` como respaldo en Unix;
    - 0.0 si no hay forma de medirla (Windows sin psutil).
    """
    if _PROC is not None:
        try:
            return _PROC.memory_info().rss / 1024 / 1024
        except Exception:
            pass
    try:
        import resource  # módulo solo Unix; en Windows lanza ImportError
        valor = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        return valor / 1024 / 1024 if sys.platform == "darwin" else valor / 1024
    except Exception:
        return 0.0


def read_excel(path, sheet, **kw):
    # keep_default_na=False: en este libro "NA"/"N/A" son VALORES de negocio
    # (54.392 celdas); tratarlos como nulos rompía los JOIN y los VLOOKUP.
    kwargs = {"keep_default_na": False}
    if ENGINE:
        kwargs["engine"] = ENGINE
    kwargs.update(kw)
    return pd.read_excel(path, sheet_name=sheet, header=0, **kwargs)


def col(df: pd.DataFrame, letra: str) -> str:
    """Nombre de columna pandas que ocupa la letra Excel indicada."""
    idx = column_index_from_string(letra) - 1
    return df.columns[idx]


def normalizar_col(nombre) -> str:
    """Clave de comparación de encabezados: sin mayúsculas, espacios ni acentos."""
    s = str(nombre).strip().lower().replace(" ", "").replace("_", "").replace(".", "")
    for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("ñ", "n")):
        s = s.replace(a, b)
    return s


def excel_weeknum_sunday(fechas: pd.Series) -> pd.Series:
    """WEEKNUM(fecha) de Excel con sistema 1 (semana empieza en domingo),
    vectorizado. Devuelve Int64 con NA donde no hay fecha."""
    valid = fechas.notna()
    out = pd.Series(pd.NA, index=fechas.index, dtype="Int64")
    if not valid.any():
        return out
    f = fechas[valid]
    # Semana Excel (sistema 1, domingo). Se evita construir el 1-ene como
    # Timestamp: hay fechas corruptas fuera de rango ns (año 2525).
    wd = f.dt.weekday                      # lunes=0 ... domingo=6
    doy = f.dt.dayofyear
    jan1_wd = (wd - (doy - 1)) % 7         # día de la semana del 1 de enero
    offset = (jan1_wd + 1) % 7             # días desde el domingo hasta el 1-ene
    week = ((doy - 1 + offset) // 7 + 1).astype("Int64")
    out.loc[valid] = week
    return out


def agregar_fechas(df: pd.DataFrame, cols: list) -> pd.Series:
    """Equivalente a IFERROR(TEXT(LARGE(rango,1),"DD/MM/YYYY"), JOIN(" - ",rango)).

    - Si el rango tiene al menos una fecha/numérico: la mayor, como DD/MM/YYYY.
    - Si no: se concatenan los valores no vacíos con ' - ' (fallback de Excel).

    El máximo se calcula en MICROSEGUNDOS columna a columna: DataFrame.max(axis=1)
    sobre fechas hace segfault en pandas 3.0.6 y, además, hay fechas del año 2525
    que no caben en nanosegundos.
    """
    mn = np.iinfo("int64").min
    best = np.full(len(df), mn, dtype="int64")
    for c in cols:
        s = pd.to_datetime(df[c], errors="coerce")
        v = s.astype("datetime64[us]").astype("int64").to_numpy()
        np.maximum(best, v, out=best)

    maxf = pd.to_datetime(pd.Series(best, index=df.index).where(best != mn),
                          unit="us", errors="coerce")
    out = maxf.dt.strftime("%d/%m/%Y")
    sin_fecha = maxf.isna()
    if sin_fecha.any():
        filas = df.loc[sin_fecha, cols]

        def _join(row):
            vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip() != ""]
            return " - ".join(vals)

        out.loc[sin_fecha] = filas.apply(_join, axis=1)
    return out.fillna("")


# ----------------------------------------------------------------------
# Recálculo vectorizado de NUEVO SEGUIMIENTO
# ----------------------------------------------------------------------
def recalcular_nuevo_seguimiento(df: pd.DataFrame,
                                 df_vj: pd.DataFrame,
                                 df_ez: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    vj_key = df_vj.columns[column_index_from_string("A") - 1]  # 'Clave'

    # --- 1) VLOOKUP a VOLCADO JIRA: destino <- origen (letra) -------------
    vlookups = {
        "M": "V",   # WO NCR1        -> NCR-1
        "N": "B",   # Estado         -> STATUS
        "S": "I",   # Hot Swap       -> HOT SWAP
        "EA": "J",  # Vendor         -> VENDOR
        "EG": "U",  # PIR            -> PIR
        "EI": "S",  # TX Ready       -> TX Ready
    }
    for dst, src in vlookups.items():
        col_src = df_vj.columns[column_index_from_string(src) - 1]
        lookup = df_vj.drop_duplicates(subset=[vj_key]).set_index(vj_key)[col_src]
        df[col(df, dst)] = df[col(df, "B")].map(lookup)

    # --- 2) VLOOKUP a Encendidos ZR (CT, CU) ------------------------------
    ez_key = df_ez.columns[column_index_from_string("A") - 1]
    for dst, idx in (("CT", 12), ("CU", 13)):
        col_src = df_ez.columns[idx]
        lookup = df_ez.drop_duplicates(subset=[ez_key]).set_index(ez_key)[col_src]
        df[col(df, dst)] = df[col(df, "B")].map(lookup)

    # --- 3) WEEKNUM(CY), WEEKNUM(DC), WEEKNUM(DJ) --------------------------
    for dst, src in (("CZ", "CY"), ("DD", "DC"), ("DK", "DJ")):
        fechas = pd.to_datetime(df[col(df, src)], errors="coerce")
        semana = excel_weeknum_sunday(fechas)
        semana = semana.where(fechas.notna(), 0).fillna(0).astype(int)
        df[col(df, dst)] = semana

    # --- 4) Estado diseño radio (DM) ---------------------------------------
    cs = df[col(df, "CS")]
    cw = df[col(df, "CW")]
    cond_cs = cs.notna() & (cs.astype(str).str.strip() != "")
    cond_cw = cw.notna() & (cw.astype(str).str.strip() != "")
    df[col(df, "DM")] = np.where(
        cond_cs,
        np.where(cond_cw, "LCC DISEÑO ACABADO", "PENDIENTE DISEÑO RADIO"),
        "",
    )

    # --- 5) Fecha máxima por bloques (DP..DU) ------------------------------
    bloques = {
        "DP": ("AD", "AN"),  # EMF VODAFONE
        "DQ": ("AO", "AX"),  # EMF ORANGE
        "DR": ("AZ", "BJ"),  # MRE VODAFONE
        "DS": ("BK", "BT"),  # MRE ORANGE
        "DT": ("BV", "CG"),  # PSC VODAFONE
        "DU": ("CH", "CQ"),  # PSC ORANGE
    }
    for dst, (ini, fin) in bloques.items():
        i0 = df.columns.get_loc(col(df, ini))
        i1 = df.columns.get_loc(col(df, fin))
        df[col(df, dst)] = agregar_fechas(df, df.columns[i0:i1 + 1].tolist())

    return df


# ----------------------------------------------------------------------
# Copia por encabezados + detección dinámica de la fila de encabezados
# ----------------------------------------------------------------------
COLS_VJ = ["Clave", "Estado", "Aprobación Plano", "Prevision RFOA", "Site Ready",
           "Acta Towerco Aprobada", "Aprobación Acta", "PrePO ", "Hot Swap", "Vendor",
           "WO Config BH", "Proyecto & Subproyecto", "Proyectos Radio", "OSP Codigo SR",
           "Cluster despliegue", "Ingenieria", "Ingeniería Zona", "Ingeniería Infra",
           "TX Ready", "Tx Disponible", "PIR", "WO NCR1"]


def detectar_encabezado(path, sheet, max_scan: int = 30, wb=None):
    """Localiza la fila de encabezados de una hoja.

    Devuelve (fila 1-based, lista de encabezados no vacíos). Busca la primera
    fila, entre las `max_scan` primeras, con al menos 5 celdas y que contenga
    'clave'; si no la encuentra, la primera con al menos 5 valores.

    Si se pasa `wb` (libro openpyxl ya abierto en read_only), se reutiliza y no
    se cierra: así se escanean varias hojas del mismo archivo abriéndolo una vez.
    """
    propio = wb is None
    if propio:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if isinstance(sheet, int):
            sheet = wb.sheetnames[sheet]
        ws = wb[sheet]
        filas = []
        for r, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan,
                                             values_only=True), 1):
            # Se conserva el texto EXACTO del encabezado (p. ej. "PrePO " con
            # espacio final): usecols de pandas compara literalmente.
            vals = [str(v) for v in row
                    if v is not None and str(v).strip() != ""]
            filas.append((r, vals))
        for r, vals in filas:
            if len(vals) >= 5 and any("clave" in v.strip().lower() for v in vals):
                return r, vals
        for r, vals in filas:
            if len(vals) >= 5:
                return r, vals
        return 1, (filas[0][1] if filas else [])
    finally:
        if propio:
            wb.close()


def cobertura_hoja(wb, nombre: str, deseadas, max_scan: int = 30):
    """Nº de columnas pedidas presentes en la hoja + fila de encabezado."""
    hr, headers = detectar_encabezado(None, nombre, max_scan=max_scan, wb=wb)
    presentes = {normalizar_col(h) for h in headers}
    return sum(1 for d in deseadas if normalizar_col(d) in presentes), hr


def elegir_hoja(wb, deseadas, preferidas=()):
    """Elige la hoja con MÁS columnas pedidas (empate: orden de `preferidas`).

    Así no dependemos de que la hoja se llame 'Hoja 2': si en otro equipo la
    llaman 'Hoja 1', 'Datos' o cambia el orden, se selecciona la que de verdad
    contiene los encabezados necesarios.
    """
    nombres = list(wb.sheetnames)
    orden = [p for p in preferidas if p in nombres]
    orden += [n for n in nombres if n not in orden]
    mejor = (None, -1, 1)
    for nombre in orden:
        try:
            cov, hr = cobertura_hoja(wb, nombre, deseadas)
        except Exception:
            continue
        if cov > mejor[1]:
            mejor = (nombre, cov, hr)
    return mejor


def elegir_hoja_archivo(path, deseadas, preferidas=()):
    """Igual que elegir_hoja, abriendo y cerrando el libro."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return elegir_hoja(wb, deseadas, preferidas)
    finally:
        wb.close()


def leer_por_encabezados(path, sheet, deseadas, wb=None) -> pd.DataFrame:
    """Lee una hoja detectando la fila de encabezados y quedándose solo con las
    columnas pedidas, emparejadas por NOMBRE (nunca por posición)."""
    hr, headers = detectar_encabezado(path, sheet, wb=wb)
    mapa = {normalizar_col(h): h for h in headers}
    pares = [(d, mapa[normalizar_col(d)]) for d in deseadas
             if normalizar_col(d) in mapa]
    faltan = [d for d in deseadas if normalizar_col(d) not in mapa]
    if not pares:
        raise ValueError(f"{Path(path).name}/{sheet}: no se encontró ninguna columna pedida")

    df = pd.read_excel(path, sheet_name=sheet, header=hr - 1,
                       usecols=[raw for _, raw in pares],
                       keep_default_na=False, engine=ENGINE)
    df = df.rename(columns={raw: canon for canon, raw in pares})
    print(f"    {Path(path).name} / {sheet}: encabezado en fila {hr}, "
          f"{len(df.columns)}/{len(deseadas)} columnas"
          + (f" (faltan: {', '.join(faltan)})" if faltan else ""))
    return df


def copiar_por_encabezados(df_origen: pd.DataFrame, df_destino: pd.DataFrame,
                           clave_origen: str = "Clave",
                           clave_destino: str = "Clave",
                           verbose: bool = False) -> pd.DataFrame:
    """Copia columnas de origen a destino emparejando por NOMBRE de encabezado.

    Manda el destino: conserva sus filas y su orden, y actualiza cada columna
    común con el valor del origen allí donde el origen tiene dato.
    """
    cols_o = {normalizar_col(c): c for c in df_origen.columns}
    cols_d = {normalizar_col(c): c for c in df_destino.columns}
    comunes = [k for k in cols_o if k in cols_d]

    kn_o, kn_d = normalizar_col(clave_origen), normalizar_col(clave_destino)
    if kn_o not in cols_o or kn_d not in cols_d:
        raise KeyError(f"Clave '{clave_origen}'/'{clave_destino}' no encontrada")

    col_o, col_d = cols_o[kn_o], cols_d[kn_d]
    a_traer = [cols_o[k] for k in comunes if k != kn_o] + [col_o]
    origen = df_origen[a_traer].drop_duplicates(subset=[col_o])

    # Clave temporal para que pandas no colapse/elimine la clave del destino
    # cuando origen y destino usan el mismo nombre de columna.
    tmp = "__clave_origen__"
    origen = origen.rename(columns={col_o: tmp})
    merged = df_destino.merge(origen, left_on=col_d, right_on=tmp,
                              how="left", suffixes=("", "__new"))
    merged = merged.drop(columns=[tmp])

    n = 0
    for k in comunes:
        if k == kn_o:
            continue
        nuevo = f"{cols_o[k]}__new"
        if nuevo in merged.columns:
            merged[cols_d[k]] = merged[nuevo].combine_first(merged[cols_d[k]])
            n += 1
    merged = merged.drop(columns=[c for c in merged.columns if c.endswith("__new")])
    if verbose:
        print(f"    {n} columnas copiadas por encabezado")
    return merged


def refrescar_volcado(df_vj: pd.DataFrame) -> pd.DataFrame:
    """Pasos 2-4 del flujo original, en memoria y con encabezados dinámicos.

    Paso 2: JIRA con DR  ->  Hoja sin título (Hoja 2)
    Paso 3: Hoja sin título -> IR_VDF (consolidado en memoria, sin escribir
            un JIRA.xlsx intermedio en disco)
    Paso 4: IR_VDF -> VOLCADO JIRA
    """
    hoja_grande = HOJA_SIN_TITULO

    # Un solo openpyxl por libro para escanear las cabeceras de sus hojas.
    wb_hoja = openpyxl.load_workbook(hoja_grande, read_only=True, data_only=True)
    try:
        print("\n[refresh] PASO 2: JIRA con DR -> Hoja sin título (por encabezados)")
        t = time.time()
        # JIRA con DR: elegir la hoja con más columnas válidas (no asumir la 1ª)
        hoja_j, cov_j, _ = elegir_hoja_archivo(JIRA_DR, COLS_VJ)
        if not hoja_j or cov_j < 5:
            raise ValueError(f"{JIRA_DR.name}: ninguna hoja tiene los encabezados esperados")
        print(f"    JIRA con DR: hoja '{hoja_j}' ({cov_j}/{len(COLS_VJ)} columnas)")
        dj = leer_por_encabezados(JIRA_DR, hoja_j, COLS_VJ)

        # Hoja sin título: preferir 'Hoja 2', pero decidir por cobertura real
        hoja_h, cov_h, _ = elegir_hoja(wb_hoja, COLS_VJ, preferidas=("Hoja 2",))
        if not hoja_h or cov_h < 5:
            raise ValueError(f"{hoja_grande.name}: ninguna hoja tiene los encabezados esperados")
        print(f"    Hoja sin título: hoja '{hoja_h}' ({cov_h}/{len(COLS_VJ)} columnas)")
        dh = leer_por_encabezados(hoja_grande, hoja_h, COLS_VJ, wb=wb_hoja)
        dh = copiar_por_encabezados(dj, dh, "Clave", "Clave", verbose=True)
        del dj
        gc.collect()
        print(f"    paso 2: {time.time() - t:.2f}s | pico {rss_mb():.0f} MB")

        print("[refresh] PASO 3: consolidar Hoja sin título -> IR_VDF (en memoria)")
        t = time.time()
        partes = [dh]
        for hoja in ("Hoja 1", "ing pdte"):
            if hoja == hoja_h or hoja not in wb_hoja.sheetnames:
                continue
            try:
                d = leer_por_encabezados(hoja_grande, hoja, COLS_VJ, wb=wb_hoja)
                if not d.empty:
                    partes.append(d)
            except Exception as e:
                print(f"    [AVISO] {hoja}: {e}")
    finally:
        wb_hoja.close()

    ir_vdf = pd.concat(partes, ignore_index=True, sort=False)
    ir_vdf = ir_vdf.dropna(subset=["Clave"]).drop_duplicates(subset=["Clave"])
    del dh, partes
    gc.collect()
    print(f"    IR_VDF: {ir_vdf.shape[0]} claves ({time.time() - t:.2f}s | "
          f"pico {rss_mb():.0f} MB)")

    print("[refresh] PASO 4: IR_VDF -> VOLCADO JIRA (por encabezados)")
    t = time.time()
    antes = df_vj.copy()
    df_vj = copiar_por_encabezados(ir_vdf, df_vj, "Clave", "Clave", verbose=True)
    del ir_vdf
    gc.collect()

    comunes = [c for c in COLS_VJ if c in antes.columns and c in df_vj.columns
               and c != "Clave"]
    cambios = sum(int((antes[c].astype(str).str.strip()
                       != df_vj[c].astype(str).str.strip()).sum())
                  for c in comunes)
    del antes
    print(f"    VOLCADO JIRA: {df_vj.shape[0]} filas, {cambios:,} celdas actualizadas "
          f"({time.time() - t:.2f}s | pico {rss_mb():.0f} MB)")
    return df_vj


def exportar_jira() -> None:
    """Genera JIRA.xlsx (hoja IR_VDF) para la Fase 5.

    Esquema: el del JIRA.xlsx existente (261 columnas) o, si no existe, el de
    la hoja de 'Hoja sin título' con más columnas.
    Datos: base = 'Hoja sin título' (todas las filas); después se actualizan
    las columnas coincidentes con JIRA con DR (que puede ser un subconjunto).
    Estructura: fila1 'VODAFONE', fila2 'IR_COMPLETO (Fiber2U)', fila3 vacía,
    fila4 cabeceras, datos desde fila5 (como espera la macro copiar_jira).
    """
    destino = BASE_DIR / "JIRA.xlsx"
    t = time.time()

    # 1) Esquema canónico
    if destino.is_file():
        try:
            hr_c, canon = detectar_encabezado(destino, "IR_VDF")
        except Exception:
            hr_c, canon = 1, []
    else:
        hr_c, canon = 1, []
    if len(canon) < 5:
        wb_h = openpyxl.load_workbook(HOJA_SIN_TITULO, read_only=True, data_only=True)
        try:
            hoja_h, _, _ = elegir_hoja(wb_h, COLS_VJ, preferidas=("Hoja 2",))
            _, canon = detectar_encabezado(None, hoja_h, wb=wb_h)
        finally:
            wb_h.close()
    canon = list(dict.fromkeys(canon))
    if "Clave" not in canon:
        raise ValueError("El esquema de JIRA no contiene la columna 'Clave'")
    print(f"    esquema IR_VDF: {len(canon)} columnas")

    # 2) Base: Hoja sin título (todas las filas del consolidado)
    wb_h = openpyxl.load_workbook(HOJA_SIN_TITULO, read_only=True, data_only=True)
    try:
        hoja_h, _, _ = elegir_hoja(wb_h, canon, preferidas=("Hoja 2",))
        hr_h, headers_h = detectar_encabezado(None, hoja_h, wb=wb_h)
    finally:
        wb_h.close()
    cols_h = [c for c in canon if c in headers_h]
    dh = pd.read_excel(HOJA_SIN_TITULO, sheet_name=hoja_h, header=hr_h - 1,
                       usecols=cols_h, keep_default_na=False, engine=ENGINE)
    print(f"    base Hoja sin título: hoja '{hoja_h}', {len(dh):,} filas, "
          f"{len(cols_h)} columnas")

    # 3) JIRA con DR: actualiza (no elimina filas) las columnas coincidentes
    hoja_j, _, _ = elegir_hoja_archivo(JIRA_DR, canon)
    hr_j, headers_j = detectar_encabezado(JIRA_DR, hoja_j)
    cols_j = [c for c in canon if c in headers_j]
    dj = pd.read_excel(JIRA_DR, sheet_name=hoja_j, header=hr_j - 1,
                       usecols=cols_j, keep_default_na=False, engine=ENGINE)
    print(f"    JIRA con DR: hoja '{hoja_j}', {len(dj):,} filas")

    clave_h = dh["Clave"].astype(str).str.strip()
    dh["Clave"] = clave_h
    dj["Clave"] = dj["Clave"].astype(str).str.strip()
    dj = (dj[dj["Clave"] != ""]
          .drop_duplicates(subset=["Clave"], keep="first")
          .set_index("Clave"))

    actualizadas = 0
    for col in [c for c in cols_j if c != "Clave"]:
        if col not in dh.columns:
            dh[col] = None
        mapeado = dj[col].reindex(clave_h)
        mascara = mapeado.notna() & (mapeado.astype(object) != "")
        if not mascara.any():
            continue
        dh[col] = np.where(mascara.to_numpy(), mapeado.to_numpy(), dh[col].to_numpy())
        actualizadas += 1
    print(f"    columnas actualizadas desde JIRA con DR: {actualizadas}")

    # 4) Orden canónico y limpieza
    for c in canon:
        if c not in dh.columns:
            dh[c] = None
    df = dh[canon]
    df = df[df["Clave"] != ""].drop_duplicates(subset=["Clave"], keep="first")
    print(f"    IR_VDF final: {len(df):,} claves")

    # 5) Escritura en streaming (solo celdas con valor)
    with xlsxwriter.Workbook(str(destino), {"constant_memory": True}) as wb:
        ws = wb.add_worksheet("IR_VDF")
        ws.write_string(0, 0, "VODAFONE")
        ws.write_string(1, 0, "IR_COMPLETO (Fiber2U)")
        for j, h in enumerate(canon):
            ws.write_string(3, j, str(h))
        fmt_fecha = wb.add_format({"num_format": "dd/mm/yyyy"})
        datos = df.to_numpy()
        for i, fila in enumerate(datos, start=4):
            for j, v in enumerate(fila):
                if v is None or v is pd.NaT:
                    continue
                if isinstance(v, float) and v != v:
                    continue
                if isinstance(v, pd.Timestamp):
                    v = v.to_pydatetime()
                if isinstance(v, datetime):
                    ws.write_datetime(i, j, v, fmt_fecha)
                elif isinstance(v, date):
                    ws.write_datetime(i, j, datetime(v.year, v.month, v.day), fmt_fecha)
                elif isinstance(v, (int, float, np.integer, np.floating)):
                    ws.write_number(i, j, float(v))
                else:
                    sval = str(v)
                    if sval != "":
                        ws.write_string(i, j, sval)
    print(f"    JIRA.xlsx escrito ({destino.stat().st_size/1e6:.1f} MB, "
          f"{time.time()-t:.1f}s)")



# ----------------------------------------------------------------------
# Validación cruzada
# ----------------------------------------------------------------------
RECALC_COLS = ["M", "N", "S", "CT", "CU", "CZ", "DD", "DK", "DM",
               "DP", "DQ", "DR", "DS", "DT", "DU", "EA", "EG", "EI"]


def _norm(v):
    if pd.isna(v):
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def validar(orig: pd.DataFrame, nuevo: pd.DataFrame) -> None:
    print("\n" + "=" * 70)
    print("  VALIDACIÓN CRUZADA  (generado vs valores cacheados del original)")
    print("=" * 70)

    recalc = {col(orig, l) for l in RECALC_COLS}
    total_ok = total_cmp = 0

    print("\n-- Columnas no recalculadas (deben coincidir al 100%) --")
    malas = []
    for c in orig.columns:
        if c in recalc:
            continue
        a = orig[c].map(_norm)
        b = nuevo[c].map(_norm)
        eq = (a == b) | (a.isna() & b.isna())
        ok, n = int(eq.sum()), len(eq)
        total_ok += ok
        total_cmp += n
        if ok != n:
            malas.append((c, ok, n))
    if malas:
        for c, ok, n in malas[:10]:
            print(f"   [FALLO] {c}: {ok}/{n}")
    else:
        print("   [OK] Todas las columnas de datos coinciden exactamente")

    print("\n-- Columnas recalculadas (vs valor cacheado de Excel) --")
    patron_fecha = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    for letra in RECALC_COLS:
        c = col(orig, letra)
        a, b = orig[c], nuevo[c]
        if letra in ("CZ", "DD", "DK"):
            mask = pd.to_numeric(a, errors="coerce").notna()
        elif letra in ("DP", "DQ", "DR", "DS", "DT", "DU"):
            mask = a.map(lambda v: bool(patron_fecha.match(str(v))) if pd.notna(v) else False)
        else:
            mask = a.notna()
        if not mask.any():
            print(f"   - {letra} ({c}): sin filas comparables")
            continue
        eq = a[mask].map(_norm) == b[mask].map(_norm)
        ok, n = int(eq.sum()), int(mask.sum())
        total_ok += ok
        total_cmp += n
        estado = "OK" if ok == n else "!!"
        print(f"   {estado} {letra:>2} ({c[:38]:<38}): {ok:>5}/{n:<5} ({ok/n:6.1%})")

    print("\n" + "-" * 70)
    print(f"  GLOBAL: {total_ok:,} / {total_cmp:,} celdas comparadas coinciden "
          f"({total_ok / total_cmp:.2%})")
    print("-" * 70)

    rng = np.random.default_rng(42)
    idx = rng.choice(len(orig), size=min(5, len(orig)), replace=False)
    cols_m = ["B", "M", "N", "EA", "EG", "EI", "CZ", "DM", "DP", "DR"]
    print("\n-- Muestra aleatoria (generado | original) --")
    for i in idx:
        partes = []
        for l in cols_m:
            partes.append(f"{l}={_norm(nuevo.iloc[i][col(orig, l)])!s} | {_norm(orig.iloc[i][col(orig, l)])!s}")
        print(f"   fila {i}: " + " ; ".join(partes))


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main() -> None:
    t_total = time.time()
    print("=" * 70)
    print("  FASE 1 - SEGUIMIENTO E2E VODAFONE (optimizado)")
    print("=" * 70)
    print(f"Motor de lectura: {ENGINE or 'openpyxl'} | refresh={REFRESCAR} | filas={FILAS or 'todas'}")

    t0 = time.time()
    df_ns = read_excel(FICHERO, HOJA_NS, nrows=FILAS)
    print(f"[1] NUEVO SEGUIMIENTO: {df_ns.shape[0]} x {df_ns.shape[1]}  ({time.time() - t0:.2f}s)")

    t0 = time.time()
    df_vj = read_excel(FICHERO, HOJA_VJ, usecols="A:V").dropna(how="all")
    print(f"[2] VOLCADO JIRA:      {df_vj.shape[0]} x {df_vj.shape[1]}  ({time.time() - t0:.2f}s)")

    t0 = time.time()
    df_ez = read_excel(FICHERO, HOJA_EZ).dropna(how="all")
    print(f"[3] Encendidos ZR:     {df_ez.shape[0]} x {df_ez.shape[1]}  ({time.time() - t0:.2f}s)")
    print(f"    Pico parcial: {rss_mb():.0f} MB")

    if REFRESCAR:
        df_vj = refrescar_volcado(df_vj)
        if "--no-export-jira" not in sys.argv:
            print("[refresh] Exportando JIRA.xlsx (IR_VDF) para la Fase 5...")
            try:
                exportar_jira()
            except OSError as e:
                print(f"[AVISO] No se pudo escribir JIRA.xlsx: {e}")
                print("        Ciérralo si está abierto en Excel; el resto continúa.")

    t0 = time.time()
    orig = df_ns.copy()
    df_out = recalcular_nuevo_seguimiento(df_ns, df_vj, df_ez)
    print(f"[4] Recálculo vectorizado: {time.time() - t0:.2f}s | pico {rss_mb():.0f} MB")

    validar(orig, df_out)

    t0 = time.time()
    salida = OUT_DIR / "seguimiento.xlsx"
    # Escritura en streaming real (escritura secuencial, memoria constante).
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(HOJA_NS)
    ws.append(list(df_out.columns))
    limpio = df_out.astype(object).where(pd.notna(df_out), None)
    for fila in limpio.itertuples(index=False, name=None):
        ws.append(fila)
    wb.save(salida)
    print(f"\n[5] Escrito {salida.name}: {salida.stat().st_size / 1e6:.2f} MB "
          f"({time.time() - t0:.2f}s)")

    print("\n" + "=" * 70)
    print(f"  TOTAL: {time.time() - t_total:.2f}s | PICO DE MEMORIA: {rss_mb():.0f} MB")
    print("=" * 70)


if __name__ == "__main__":
    main()