#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cálculo de PREVISIONES JUAN sin Excel, por bloques de filas con pycel.

pycel no soporta LET ni XLOOKUP de forma nativa en esta versión, así que:
  - LET se expande a su expresión.
  - XLOOKUP se reescribe como IFERROR(INDEX(...),MATCH(...)).
  - Los rangos completos (A:FS, A:A) se acotan a las dimensiones reales.
  - WEEKNUM y SEARCH se añaden como funciones de pycel.

El libro se divide en bloques de 50 filas para que el grafo de pycel sea
manejable; se calculan todos los bloques y se devuelven los valores.
"""

from __future__ import annotations

import ast
import datetime as dt
import re
import shutil
import tempfile
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

CHUNK = 50


# ----------------------------------------------------------------------
# Compatibilidad pycel / Python 3.15 y funciones que faltan
# ----------------------------------------------------------------------
def _preparar_pycel():
    if not hasattr(ast, "Str"):
        def _Str(s=None, **kw):
            return ast.Constant(value=s if s is not None else kw.get("s"))
        ast.Str = _Str

    import pycel.lib.date_time as dtm
    import pycel.lib.text as txt

    def weeknum(serial, return_type=1):
        if serial in ("", None, "-"):
            return ""
        if isinstance(serial, dt.datetime):
            d = serial.date()
        elif isinstance(serial, dt.date):
            d = serial
        else:
            d = (dt.datetime(1899, 12, 30) +
                 dt.timedelta(days=float(serial))).date()
        if int(return_type) == 21:
            return d.isocalendar()[1]
        jan1 = dt.date(d.year, 1, 1)
        off = (jan1.weekday() + 1) % 7
        return (d.timetuple().tm_yday + off - 1) // 7 + 1

    def search(find_text, within_text, start=1):
        try:
            i = str(within_text).lower().find(str(find_text).lower(), int(start) - 1)
            return i + 1 if i >= 0 else "#VALUE!"
        except Exception:
            return "#VALUE!"

    dtm.weeknum = weeknum
    txt.search = search

    import pycel.excelwrapper as ew
    ew.ExcelOpxWrapper.defined_names = property(lambda self: {})

    from pycel import ExcelCompiler
    return ExcelCompiler


# ----------------------------------------------------------------------
# Expansión de LET / XLOOKUP y acotado de rangos
# ----------------------------------------------------------------------
def _split_args(s: str) -> list:
    out, depth, q, start = [], 0, False, 0
    for i, ch in enumerate(s):
        if ch == '"':
            q = not q
        elif not q:
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
            elif ch == ',' and depth == 0:
                out.append(s[start:i])
                start = i + 1
    out.append(s[start:])
    return out


def _expandir_let(f: str) -> str:
    for pref in ("_xlfn.LET(", "LET("):
        while True:
            i = f.find(pref)
            if i < 0:
                break
            j, depth, q = i + len(pref), 1, False
            while j < len(f) and depth:
                if f[j] == '"':
                    q = not q
                elif not q and f[j] == '(':
                    depth += 1
                elif not q and f[j] == ')':
                    depth -= 1
                j += 1
            if depth:
                break
            args = _split_args(f[i + len(pref):j - 1])
            if len(args) < 3 or len(args) % 2 == 0:
                break
            expr = args[-1]
            for k in range(0, len(args) - 1, 2):
                name = args[k].strip()
                val = _expandir_let(args[k + 1].strip())
                expr = re.sub(r"\b" + re.escape(name) + r"\b", f"({val})", expr)
            f = f[:i] + "(" + _expandir_let(expr) + ")" + f[j:]
    return f.replace("_xlfn.XLOOKUP(", "XLOOKUP(").replace("_xlfn.", "")


def _reescribir_xlookup(f: str) -> str:
    while True:
        i = f.find("XLOOKUP(")
        if i < 0:
            break
        j, depth, q = i + len("XLOOKUP("), 1, False
        while j < len(f) and depth:
            if f[j] == '"':
                q = not q
            elif not q and f[j] == '(':
                depth += 1
            elif not q and f[j] == ')':
                depth -= 1
            j += 1
        if depth:
            break
        a = _split_args(f[i + len("XLOOKUP("):j - 1])
        if len(a) < 3:
            break
        default = a[3].strip() if len(a) > 3 else "NA()"
        rep = (f"IFERROR(INDEX({a[2].strip()},"
               f"MATCH({a[0].strip()},{a[1].strip()},0)),{default})")
        f = f[:i] + rep + f[j:]
    return f


def _acotar(f: str, limites: dict) -> str:
    for hoja, (maxr, maxc) in limites.items():
        f = re.sub(
            r"""'""" + re.escape(hoja) + r"""'!\$?([A-Z]+):\$?([A-Z]+)""",
            lambda m, h=hoja, maxr=maxr: (
                f"'{h}'!${m.group(1)}$1:${m.group(2)}${maxr}"),
            f,
        )
        f = f.replace(f"'{hoja}'!$1:$2",
                      f"'{hoja}'!$A$1:${get_column_letter(maxc)}$2")
    f = re.sub(r"drive!\$?A:\$?A",
               f"drive!$A$1:$A${limites['drive'][0]}", f)
    return f


# ----------------------------------------------------------------------
# Carga de datos y construcción de bloques
# ----------------------------------------------------------------------
class Datos:
    def __init__(self, origen: Path):
        self.origen = origen
        self.wf = openpyxl.load_workbook(origen, data_only=False)
        self.wv = openpyxl.load_workbook(origen, data_only=True)
        self.pf = self.wf["PREVISIONES JUAN"]
        self.pv = self.wv["PREVISIONES JUAN"]
        self.headers = [self.pv.cell(row=1, column=c).value for c in range(1, 88)]
        self.irs = [self.pf.cell(row=r, column=1).value
                    for r in range(2, self.pf.max_row + 1)]

        di = self.wf["DATA IN"]
        div = self.wv["DATA IN"]
        hdr = [di.cell(row=1, column=c).value for c in range(1, 262)]
        self.di_cols = [1] + [c for c in range(2, 262) if hdr[c - 1] in set(self.headers)]
        self.di_hdr = hdr
        self.di_rows = {}
        for r in range(3, di.max_row + 1):
            ir = div.cell(row=r, column=1).value
            if ir not in (None, "") and ir not in self.di_rows:
                self.di_rows[ir] = r
        self.div = div

        self.legal = {}
        for nombre in ("LEGAL OSP", "LEGAL VDF"):
            sw = self.wv[nombre]
            rows = {}
            for r in range(2, sw.max_row + 1):
                ir = sw.cell(row=r, column=1).value
                if ir not in (None, "") and ir not in rows:
                    rows[ir] = r
            self.legal[nombre] = (sw, rows)

        self.drive = self.wv["drive"]
        self.limites = {
            "DATA IN": (max(3, len(self.di_rows) + 2), len(self.di_cols)),
            "LEGAL OSP": (max(2, self.legal["LEGAL OSP"][0].max_row),
                          self.legal["LEGAL OSP"][0].max_column),
            "LEGAL VDF": (max(2, self.legal["LEGAL VDF"][0].max_row),
                          self.legal["LEGAL VDF"][0].max_column),
            "drive": (self.drive.max_row, 4),
        }

    def bloque(self, irs: list, destino: Path) -> None:
        wb = Workbook()
        wp = wb.active
        wp.title = "PREVISIONES JUAN"
        for c in range(1, 88):
            wp.cell(row=1, column=c, value=self.headers[c - 1])
        filas = [r for r in range(2, self.pf.max_row + 1)
                 if self.pf.cell(row=r, column=1).value in set(irs)]
        for r, src in enumerate(filas, start=2):
            for c in range(1, 88):
                v = self.pf.cell(row=src, column=c).value
                if isinstance(v, ArrayFormula):
                    v = v.text
                if isinstance(v, str) and v.startswith("="):
                    v = _reescribir_xlookup(_expandir_let(v))
                    v = _acotar(v, self.limites)
                wp.cell(row=r, column=c, value=v)

        wd = wb.create_sheet("DATA IN")
        for j, c in enumerate(self.di_cols, 1):
            wd.cell(row=1, column=j, value=self.di_hdr[c - 1])
            wd.cell(row=2, column=j, value=j)
        out = 3
        for ir in irs:
            r = self.di_rows.get(ir)
            if r is None:
                continue
            for j, c in enumerate(self.di_cols, 1):
                wd.cell(row=out, column=j,
                        value=self.div.cell(row=r, column=c).value)
            out += 1

        for nombre in ("LEGAL OSP", "LEGAL VDF"):
            sw, rows = self.legal[nombre]
            nw = wb.create_sheet(nombre)
            for r in range(1, sw.max_column + 1):
                nw.column_dimensions[get_column_letter(r)].width = 14
            out = 1
            for ir in irs:
                sr = rows.get(ir)
                if sr is None:
                    continue
                for c in range(1, sw.max_column + 1):
                    nw.cell(row=out, column=c, value=sw.cell(row=sr, column=c).value)
                out += 1

        nd = wb.create_sheet("drive")
        for r in range(1, self.drive.max_row + 1):
            for c in range(1, 5):
                nd.cell(row=r, column=c, value=self.drive.cell(row=r, column=c).value)
        wb.save(destino)
        wb.close()


# ----------------------------------------------------------------------
# Cálculo por bloques
# ----------------------------------------------------------------------
def calcular_valores(origen: Path, progreso=None) -> dict:
    ExcelCompiler = _preparar_pycel()
    datos = Datos(origen)
    valores = {}
    total = len(datos.irs)
    if progreso:
        progreso(0, total)

    for ini in range(0, total, CHUNK):
        irs = datos.irs[ini:ini + CHUNK]
        tmp = Path(tempfile.mkdtemp(prefix="semi_calc_")) / "bloque.xlsx"
        try:
            datos.bloque(irs, tmp)
            py = ExcelCompiler(str(tmp), cycles=True)
            for i, ir in enumerate(irs, start=2):
                for c in range(1, 88):
                    addr = f"'PREVISIONES JUAN'!{get_column_letter(c)}{i}"
                    try:
                        valores[(i, c)] = py.evaluate(addr)
                    except Exception:
                        valores[(i, c)] = None
        finally:
            shutil.rmtree(tmp.parent, ignore_errors=True)
        if progreso:
            progreso(min(ini + CHUNK, total), total)
    datos.wf.close()
    datos.wv.close()
    return valores


def convertir_a_valores(path: Path) -> None:
    """Calcula el libro y sustituye sus fórmulas por los valores obtenidos."""
    valores = calcular_valores(path)
    wb = openpyxl.load_workbook(path)
    ws = wb["PREVISIONES JUAN"]
    for (fila, col), valor in valores.items():
        ws.cell(row=fila, column=col, value=valor)
    wb.save(path)
    wb.close()