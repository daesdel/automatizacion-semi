#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FASE 2 - Legal: MinisterioPowerBi.csv -> Legal_Final.xlsx
=========================================================

Replica en Python el flujo de "Procesado legal automatico.xlsm":

  1. Leer MinisterioPowerBi.csv (30 columnas, separador coma).
  2. Limpiar fechas: quitar "T00:00:00" y convertir seriales Excel (1899-12-30).
  3. Reordenar columnas Y:AA -> Q (igual que la macro ProcesarCSV).
  4. Separar por 'Operador Responsable': VDF (Vodafone) y OSP (Orange).
  5. Construir el equivalente a Procesado_vdf / Procesado_osp:
        clave + tecnologia + EMF + AMT + PSC   (VLOOKUP por clave_tecnologia)
  6. Construir la matriz de Ministerio VDF / Ministerio_OSP:
        1 fila por clave única, 3 bloques (EMF/AMT/PSC) x 11 tecnologías.
  7. Escribir Legal_Final.xlsx con dos hojas (LEGAL_VDF, LEGAL_OSP), solo valores.

Uso:
    PYTHONPATH=/tmp/my_py python3 automatizacion/fase2_legal.py
"""

import sys
import time
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter

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


# Consola de Windows (cp1252): evita UnicodeEncodeError al imprimir símbolos.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE_DIR = Path(__file__).parent.parent / "Nueva carpeta (2)"
CSV_PATH = _arg_path("csv", BASE_DIR / "MinisterioPowerBi.csv")
OUT_DIR = _arg_path("out", BASE_DIR / "resultados")
OUT_PATH = OUT_DIR / f"{date.today():%d%m%Y}_legal ministerio.xlsx"

COL_CLAVE = "Clave de incidencia"
COL_TEC = "Tecnologia"
COL_OPERADOR = "Operador Responsable"
COL_EMF = "Fecha Entrega MINECO/EMF"
COL_AMT = "Fecha AMT"
COL_PSC = "Fecha FTP PS"

# Las 11 tecnologías de la matriz original (3 bloques: EMF, AMT, PSC)
TECNOLOGIAS = ["G900", "U900", "U2100", "LTE800", "LTE900", "LTE1800",
               "LTE2100", "LTE2600", "5G 700", "5G 3700", "5G 2100"]

FMT_FECHA = "dd/mm/yyyy"


# ----------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------
def norm_cabecera(c: str) -> str:
    """Quita corchetes/espacios y normaliza para comparar cabeceras."""
    return str(c).strip().strip("[]").strip()


def norm_texto(s) -> str:
    return str(s).strip().casefold() if s is not None else ""


def leer_csv() -> pd.DataFrame:
    df = pd.read_csv(CSV_PATH, sep=",", dtype=str, keep_default_na=False,
                     encoding="utf-8-sig")
    df.columns = [norm_cabecera(c) for c in df.columns]
    print(f"[1] CSV leído: {df.shape[0]:,} filas x {df.shape[1]} columnas")
    return df


def limpiar_fechas(df: pd.DataFrame) -> pd.DataFrame:
    """Quita la hora ISO y convierte seriales Excel a fecha real.

    - '2023-06-03T00:00:00' -> 2023-06-03
    - '45243'               -> 2024-11-15 (época Excel 1899-12-30)
    - Textos ('N/A', 'SI')  -> se conservan tal cual
    """
    df = df.copy()
    n_iso = n_serial = n_texto = 0
    for c in df.columns:
        if "fecha" not in norm_cabecera(c).casefold():
            continue
        s = df[c].astype(str).str.strip().str.replace("T00:00:00", "", regex=False)

        iso = pd.to_datetime(s, errors="coerce", format="%Y-%m-%d")
        num = pd.to_numeric(s.replace("", np.nan), errors="coerce")
        serial = pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")

        fechas = iso.where(iso.notna(), serial)
        parsed = fechas.notna()
        res = fechas.dt.date.astype(object).where(parsed, s)

        n_iso += int((iso.notna()).sum())
        n_serial += int((serial.notna()).sum())
        n_texto += int((~parsed & (s != "")).sum())
        df[c] = res
    print(f"[2] Fechas limpias: ISO={n_iso:,} | seriales={n_serial:,} | texto conservado={n_texto:,}")
    return df


def mover_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Replica ProcesarCSV: corta Y:AA (cols 25-27) y las inserta en Q (17)."""
    cols = list(df.columns)
    if len(cols) < 27:
        print("[3] Aviso: el CSV tiene menos de 27 columnas; no se reordena")
        return df
    bloque = cols[24:27]          # Y, Z, AA
    resto = cols[:24] + cols[27:]
    nuevas = resto[:16] + bloque + resto[16:]
    print(f"[3] Columnas Y:AA movidas a Q: {bloque}")
    return df[nuevas]


def separar_operador(df: pd.DataFrame):
    op = df[COL_OPERADOR].map(norm_texto)
    es_vdf = op.str.startswith("vodafone")
    es_osp = op.str.startswith("orange")
    otros = df[~(es_vdf | es_osp)]
    if len(otros):
        print(f"    Aviso: {len(otros)} filas con operador no reconocido: "
              f"{otros[COL_OPERADOR].value_counts().head().to_dict()}")
    vdf, osp = df[es_vdf], df[es_osp]
    print(f"[4] Separación: VDF={len(vdf):,} | OSP={len(osp):,}")
    return vdf, osp


def construir_procesado(df_op: pd.DataFrame, etiqueta: str) -> pd.DataFrame:
    """Equivale a Procesado_xxx: clave, tecnología, EMF, AMT, PSC + clave_tecnología."""
    sub = df_op[[COL_CLAVE, COL_TEC, COL_EMF, COL_AMT, COL_PSC]].copy()
    sub.columns = ["clave", "tecnologia", "emf", "amt", "psc"]
    sub["clave"] = sub["clave"].astype(str).str.strip()
    sub["tecnologia"] = sub["tecnologia"].astype(str).str.strip()
    sub = sub[sub["clave"] != ""]

    sub["key"] = sub["clave"] + "_" + sub["tecnologia"]
    antes = len(sub)
    sub = sub.drop_duplicates(subset=["key"], keep="first")
    print(f"    Procesado {etiqueta}: {antes:,} filas -> {len(sub):,} claves_tecnología únicas")
    return sub


def _max_fechas(bloque: pd.DataFrame) -> pd.Series:
    """MAX(fila) ignorando textos/vacíos, igual que Excel sobre fechas."""
    mn = np.iinfo("int64").min
    best = np.full(len(bloque), mn, dtype="int64")
    for c in bloque.columns:
        s = bloque[c].map(lambda v: pd.Timestamp(v) if isinstance(v, (date, datetime)) else pd.NaT)
        s = pd.to_datetime(s)
        v = s.astype("datetime64[us]").astype("int64").to_numpy()
        np.maximum(best, v, out=best)
    return pd.to_datetime(pd.Series(best, index=bloque.index).where(best != mn),
                          unit="us", errors="coerce")


def construir_matriz(sub: pd.DataFrame):
    """Matriz 1 fila por clave única, 33 celdas (11 tecnologías x EMF/AMT/PSC)
    y las 3 columnas consolidadas AI:AK = MAX de cada bloque."""
    claves = sub["clave"].drop_duplicates().tolist()
    idx = pd.Index(claves, name="clave")
    lookup = sub.set_index("key")

    bloques = {}
    for campo in ("emf", "amt", "psc"):
        serie = lookup[campo]
        datos = {}
        for tech in TECNOLOGIAS:
            claves_tec = [f"{c}_{tech}" for c in claves]
            datos[tech] = serie.reindex(claves_tec).to_numpy()
        bloques[campo] = pd.DataFrame(datos, index=idx)

    consolidados = {campo: _max_fechas(bloques[campo]) for campo in bloques}
    return idx, bloques, consolidados


# ----------------------------------------------------------------------
# Escritura
# ----------------------------------------------------------------------
def _valor(v):
    """Normaliza un valor para xlsxwriter: fecha, texto o vacío."""
    if v is None:
        return ""
    if isinstance(v, float) and np.isnan(v):
        return ""
    if isinstance(v, pd.Timestamp):
        return v.date()
    return v


def escribir_hoja(wb, nombre: str, etiqueta: str, idx: pd.Index, bloques: dict,
                  consolidados: dict, f_fecha, f_cab, f_grupo, f_clave):
    ws = wb.add_worksheet(nombre)

    # Fila 1: numeración 1..37
    for i in range(37):
        ws.write_number(0, i, i + 1, f_cab)

    # Fila 2: bloques de grupo (EMF/AMT/PSC) sobre 11 columnas cada uno
    ws.merge_range(1, 1, 1, 11, f"EMF {etiqueta}", f_grupo)
    ws.merge_range(1, 12, 1, 22, f"AMT {etiqueta}", f_grupo)
    ws.merge_range(1, 23, 1, 33, f"PSC {etiqueta}", f_grupo)

    # Fila 3: cabecera de claves + tecnologías + EMF/AMT/PSC
    ws.write_string(2, 0, f"[{COL_CLAVE}]", f_cab)
    for j, tech in enumerate(TECNOLOGIAS):
        ws.write_string(2, 1 + j, tech, f_cab)
        ws.write_string(2, 12 + j, tech, f_cab)
        ws.write_string(2, 23 + j, tech, f_cab)
    for k, etiqueta_col in enumerate(("EMF", "AMT", "PSC")):
        ws.write_string(2, 34 + k, etiqueta_col, f_cab)

    ws.freeze_panes(3, 1)
    ws.set_column(0, 0, 16)
    ws.set_column(1, 36, 11)

    # Datos desde la fila 4
    cols = [("emf", 1), ("amt", 12), ("psc", 23)]
    fila = 3
    for clave in idx:
        ws.write_string(fila, 0, str(clave), f_clave)
        for campo, base in cols:
            valores = bloques[campo].loc[clave].to_numpy()
            for j, tech in enumerate(TECNOLOGIAS):
                v = _valor(valores[j])
                if isinstance(v, date):
                    ws.write_datetime(fila, base + j,
                                      datetime(v.year, v.month, v.day), f_fecha)
                elif v != "":
                    ws.write_string(fila, base + j, str(v))
        # AI:AK = MAX de cada bloque (0 si no hay fechas, como Excel)
        for k, campo in enumerate(("emf", "amt", "psc")):
            ts = consolidados[campo].loc[clave]
            if pd.isna(ts):
                ws.write_number(fila, 34 + k, 0)
            else:
                ws.write_datetime(fila, 34 + k,
                                  datetime(ts.year, ts.month, ts.day), f_fecha)
        fila += 1
    print(f"    Hoja {nombre}: {len(idx):,} filas de datos")


def escribir_libro(resultados: dict):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(OUT_PATH), {"strings_to_numbers": False})
    f_cab = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
    f_grupo = wb.add_format({"bold": True, "align": "center", "bg_color": "#FCE4D6",
                             "border": 1})
    f_fecha = wb.add_format({"num_format": FMT_FECHA})
    f_clave = wb.add_format({"bold": False})

    for etiqueta, (idx, bloques, consolidados) in resultados.items():
        escribir_hoja(wb, f"legal_{etiqueta.lower()}", etiqueta, idx, bloques,
                      consolidados, f_fecha, f_cab, f_grupo, f_clave)
    wb.close()
    print(f"[7] Escrito: {OUT_PATH} ({OUT_PATH.stat().st_size / 1e6:.2f} MB)")


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main() -> None:
    t0 = time.time()
    print("=" * 70)
    print(f"  FASE 2 - LEGAL (MinisterioPowerBi.csv -> {OUT_PATH.name})")
    print("=" * 70)
    if not CSV_PATH.exists():
        sys.exit(f"No existe {CSV_PATH}")

    df = leer_csv()
    df = limpiar_fechas(df)
    df = mover_columnas(df)

    vdf, osp = separar_operador(df)

    resultados = {}
    for etiqueta, datos in (("VDF", vdf), ("OSP", osp)):
        print(f"\n[5] Construyendo {etiqueta}...")
        sub = construir_procesado(datos, etiqueta)
        idx, bloques, consolidados = construir_matriz(sub)

        # Tecnologías fuera de la matriz (se conservan en claves, sin columna)
        fuera = sorted(set(sub["tecnologia"]) - set(TECNOLOGIAS))
        if fuera:
            n = int(sub["tecnologia"].isin(fuera).sum())
            print(f"    Tecnologías sin columna en la matriz ({n:,} filas): {fuera}")

        # Cobertura de la matriz
        rellenas = sum(int(bloques[c].notna().sum().sum()) for c in bloques)
        print(f"    Claves únicas: {len(idx):,} | celdas con valor: {rellenas:,} "
              f"de {len(idx) * 33:,}")
        resultados[etiqueta] = (idx, bloques, consolidados)

    print(f"\n[6] Escribiendo {OUT_PATH.name}...")
    escribir_libro(resultados)

    print("\n" + "=" * 70)
    print(f"  TOTAL: {time.time() - t0:.2f}s")
    print("=" * 70)


if __name__ == "__main__":
    main()