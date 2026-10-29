#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FASE 5 - Archivo limpio: Seguimiento_Previsiones_Limpio.xlsx
=============================================================

Genera el libro con 8 hojas a partir de:

    Seguimiento_Previsiones_Limpio.xlsx (semana anterior) -> PREVISIONES JUAN
    JIRA.xlsx                                             -> DATA IN + detalle
    resultados/DDMMYYYY_legal ministerio.xlsx             -> LEGAL VDF / OSP
    resultados/seguimiento.xlsx                           -> drive
    Revision Legalizacion ...                             -> Leg_Wn
    Hoja sin título (Subtareas)                           -> AB / AJ
    (ZOOM y DETALLE_TICKETS se calculan)

Si no hay archivo anterior, genera PREVISIONES desde cero SIN fórmulas.

Uso:
    python automatizacion/fase5_previsiones.py
    python automatizacion/fase5_previsiones.py --out=RUTA
"""

import os
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path
import re as _re
import xlsxwriter.worksheet as _xw


class FormulasNormales:
    """Evita que xlsxwriter convierta XLOOKUP/SWITCH en fórmulas matriciales."""

    def __enter__(self):
        self._orig = _xw.re_dynamic_function
        _xw.re_dynamic_function = _re.compile(r"(?!x)x")
        return self

    def __exit__(self, *args):
        _xw.re_dynamic_function = self._orig


import openpyxl
import xlsxwriter
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.formula.translate import Translator

HERE = Path(__file__).resolve().parent
BASE_DIR = HERE.parent / "Nueva carpeta (2)"


def _arg_path(nombre: str, defecto: Path) -> Path:
    pref = f"--{nombre}="
    for a in sys.argv[1:]:
        if a.startswith(pref):
            return Path(a.split("=", 1)[1]).expanduser()
    return defecto


ANTERIOR = _arg_path("anterior", BASE_DIR / "Seguimiento_Previsiones_Limpio_anterior.xlsx")
HOJA = _arg_path("hoja", BASE_DIR / "Hoja de cálculo sin título.xlsx")
JIRA = _arg_path("jira", BASE_DIR / "JIRA.xlsx")
LEGAL = _arg_path("legal", BASE_DIR / "resultados" /
                  f"{date.today():%d%m%Y}_legal ministerio.xlsx")
SEGUIMIENTO = _arg_path("seguimiento", BASE_DIR / "resultados" / "seguimiento.xlsx")
REVISION = _arg_path("revision", BASE_DIR / "Revision Legalizacion SEMI SemN.xlsx")
DESTINO = _arg_path("out", BASE_DIR / "resultados" / "Seguimiento_Previsiones_Limpio.xlsx")

sys.path.insert(0, str(HERE))
import estilos_previsiones as est  # noqa: E402
import data_in_headers as dih  # noqa: E402
import previsiones_headers as ph  # noqa: E402

# (campo, patrón editable, descripción)
ESCENARIOS = [
    ("G", "*portadoras) 5G 7*", "PORTADORAS 5G700",
     'Proyectos Radio contiene "portadoras) 5G 7"'),
    ("G", "*portadoras) 5g 37*", "PORTADORAS 5G3700",
     'Proyectos Radio contiene "portadoras) 5g 37"'),
    ("G", "*portadoras*18*", "PORTADORAS 5G1800",
     'Proyectos Radio contiene "portadoras" y "18"'),
    ("G", "*4g*", "PORTADORAS 4G", 'Proyectos Radio contiene "4g"'),
    ("E", "*upgrade*", "CAPACIDAD UPGRADE",
     'Proyecto & Subproyecto contiene "upgrade"'),
    ("E", "*jumping*", "JUMPING", 'Proyecto & Subproyecto contiene "jumping"'),
]


def _coincide(patron: str, texto: str) -> bool:
    """Comodín simple: *a*b* -> todas las partes presentes."""
    partes = [p for p in patron.strip("*").lower().split("*") if p]
    texto = (texto or "").lower()
    return all(p in texto for p in partes)


# ----------------------------------------------------------------------
# Escritura de valores
# ----------------------------------------------------------------------
def volcar(ws, fila: int, col: int, valor, fmt=None, fmt_fecha=None) -> None:
    if valor is None or valor == "":
        return
    if isinstance(valor, ArrayFormula):
        ws.write_array_formula(fila, col, fila, col, valor.text, fmt)
    elif isinstance(valor, str) and valor.startswith("="):
        ws.write_formula(fila, col, valor, fmt)
    elif isinstance(valor, datetime):
        ws.write_datetime(fila, col, valor, fmt or fmt_fecha)
    elif isinstance(valor, date):
        ws.write_datetime(fila, col, datetime(valor.year, valor.month, valor.day),
                          fmt or fmt_fecha)
    elif isinstance(valor, bool):
        ws.write_boolean(fila, col, valor, fmt)
    elif isinstance(valor, (int, float)):
        ws.write_number(fila, col, valor, fmt)
    else:
        ws.write_string(fila, col, str(valor), fmt)


# ----------------------------------------------------------------------
# PREVISIONES JUAN (base: archivo de la semana anterior)
# ----------------------------------------------------------------------
MANUAL = {1: "ir", 8: "h", 28: "ab", 36: "aj"}


def leer_anterior(path: Path):
    """Devuelve (cabeceras, filas) de PREVISIONES JUAN del archivo anterior."""
    if not path or not path.is_file():
        return [], []
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    if "PREVISIONES JUAN" not in wb.sheetnames:
        wb.close()
        return [], []
    ws = wb["PREVISIONES JUAN"]
    it = ws.iter_rows(values_only=True)
    try:
        hdr = list(next(it))
    except StopIteration:
        wb.close()
        return [], []
    filas = []
    for fila in it:
        fila = list(fila)
        if len(fila) < 87:
            fila += [None] * (87 - len(fila))
        filas.append(fila[:87])
    wb.close()
    return hdr, filas


def leer_jira_map() -> dict:
    """IR -> datos de JIRA necesarios para H, AJ y DETALLE."""
    wb = openpyxl.load_workbook(JIRA, read_only=True, data_only=True)
    ws = wb["IR_VDF"]
    hdr = list(next(ws.iter_rows(min_row=4, max_row=4, values_only=True)))
    idx = {h: i for i, h in enumerate(hdr) if h}

    def val(fila, nombre):
        i = idx.get(nombre)
        return fila[i] if i is not None and i < len(fila) else None

    datos = {}
    for fila in ws.iter_rows(min_row=5, values_only=True):
        clave = val(fila, "Clave")
        if clave in (None, ""):
            continue
        datos[str(clave).strip()] = {
            "e": val(fila, "Proyecto & Subproyecto"),
            "g": val(fila, "Proyectos Radio"),
            "at": val(fila, "OSP Tecnología IR"),
            "t": val(fila, "Fecha PO"),
            "av": val(fila, "Prevision RFOA"),
            "au": val(fila, "Real RFOA"),
            "estado": val(fila, "Estado"),
        }
    wb.close()
    return datos


def leer_subtareas():
    """Dos mapas IR -> valor: Tx Single RAN y ACTIVACION OR."""
    wb = openpyxl.load_workbook(HOJA, read_only=True, data_only=True)
    if "Subtareas" not in wb.sheetnames:
        wb.close()
        raise RuntimeError("No existe la hoja Subtareas en Hoja sin título.")
    ws = wb["Subtareas"]
    tx, act = {}, {}
    for fila in ws.iter_rows(min_row=2, values_only=True):
        ir = str(fila[1] or "").strip()
        tipo = str(fila[6] or "").strip()
        valor = fila[5]
        if not ir:
            continue
        if tipo == "Tx Single RAN" and ir not in tx:
            tx[ir] = valor
        elif tipo == "ACTIVACION OR" and ir not in act:
            act[ir] = valor
    wb.close()
    return tx, act


def construir_irs(previas, drive_rows) -> list:
    prev = [str(f[0]).strip() for f in previas if f and f[0] not in (None, "")]
    drive_irs = [str(f[0]).strip() for f in drive_rows if f and f[0] not in (None, "")]
    nuevos = []
    if prev:
        idx = None
        if prev[-1] in drive_irs:
            idx = drive_irs.index(prev[-1])
        else:
            for ir in reversed(prev):
                if ir in drive_irs:
                    idx = drive_irs.index(ir)
                    break
        if idx is not None:
            for ir in drive_irs[idx + 1:]:
                if ir and ir not in prev and ir not in nuevos:
                    nuevos.append(ir)
    else:
        nuevos = list(dict.fromkeys(drive_irs))
    return prev + nuevos


def calcular_manual(ir: str, jmap: dict, tx: dict, act: dict):
    d = jmap.get(ir, {})
    e = str(d.get("e") or "")
    g = str(d.get("g") or "")
    nombres = []
    for campo, patron, nombre, _ in ESCENARIOS:
        texto = g if campo == "G" else e
        if _coincide(patron, texto) and nombre not in nombres:
            nombres.append(nombre)
    h = " + ".join(nombres) if nombres else e

    ab = tx.get(ir, "Sin IR")
    at = str(d.get("at") or "").strip()
    if at in ("", "-", "N/A", "n/a"):
        aj = "N/A"
    else:
        aj = act.get(ir, "Sin IR")
    return h, ab, aj


def escribir_previsiones(wb, headers, previas, irs, calc) -> None:
    ws = wb.add_worksheet("PREVISIONES JUAN")
    for col, ancho in est.ANCHOS.items():
        ws.set_column(col - 1, col - 1, ancho)
    ws.set_row(0, est.ALTO_FILA1)
    ws.freeze_panes(*est.FREEZE)
    cab, dat = est.crear_formatos(wb)

    hdr = list(headers) if len(headers) >= 87 else list(ph.HEADERS)
    for c, h in enumerate(hdr[:87], start=1):
        if h not in (None, ""):
            ws.write_string(0, c - 1, str(h), cab.get(c))

    last_row = len(previas) + 1  # fila Excel de la última fila anterior
    plantillas = {}
    if previas:
        ultima = previas[-1]
        for c in range(1, 88):
            v = ultima[c - 1] if c - 1 < len(ultima) else None
            if isinstance(v, ArrayFormula) or (isinstance(v, str) and v.startswith("=")):
                plantillas[c] = v

    filas = 1
    for i, ir in enumerate(irs, start=2):
        src = previas[i - 2] if i - 2 < len(previas) else None
        for c in range(1, 88):
            fmt = dat.get(c)
            if c == 1:
                ws.write_string(i - 1, 0, ir, fmt)
            elif c in MANUAL:
                h, ab, aj = calc.get(ir, ("", "Sin IR", "N/A"))
                valor = {"h": h, "ab": ab, "aj": aj}[MANUAL[c]]
                volcar(ws, i - 1, c - 1, valor, fmt)
            elif src is not None:
                v = src[c - 1] if c - 1 < len(src) else None
                if v != "AUTO":
                    volcar(ws, i - 1, c - 1, v, fmt)
            elif c in plantillas:
                v = plantillas[c]
                txt = v.text if isinstance(v, ArrayFormula) else v
                col = get_column_letter(c)
                try:
                    txt = Translator(txt, origin=f"{col}{last_row}").translate_formula(
                        f"{col}{i}")
                except Exception:
                    pass
                if isinstance(v, ArrayFormula):
                    ws.write_array_formula(i - 1, c - 1, i - 1, c - 1, txt, fmt)
                else:
                    ws.write_formula(i - 1, c - 1, txt, fmt)
        filas += 1

    verde = wb.add_format({"bg_color": "#92D050"})
    rojo = wb.add_format({"bg_color": "#FF0000"})
    gris = wb.add_format({"bg_color": "#8D8B8B"})
    reglas = [("Aceptado Zona", verde), ("RIP", gris), ("Doc & Inv Radio", verde),
              ("KPIS OK", verde), ("ON AIR", verde), ("RFOA", verde),
              ("Parada de Reloj", rojo)]
    for rango, ref in [(f"A2:BH{filas}", 2), ("AJ1314", 1314), ("AJ3374:AJ3376", 3374)]:
        for texto, fmt in reglas:
            ws.conditional_format(rango, {
                "type": "formula",
                "criteria": f'=$N{ref}="{texto}"',
                "format": fmt,
            })
    nuevas = max(0, len(irs) - len(previas))
    print(f"    PREVISIONES JUAN: {len(irs):,} filas "
          f"({len(previas):,} leídas del anterior + {nuevas:,} nuevas)")


# ----------------------------------------------------------------------
# DATA IN: cabeceras del modulo + datos de JIRA
# ----------------------------------------------------------------------
def escribir_data_in(wb, ws_jira, fmt_fecha, fmt_cab) -> None:
    ws = wb.add_worksheet("DATA IN")
    ws.set_column(0, 0, 18)
    ws.set_column(1, 260, 12)

    hdr = dih.HEADERS
    idx = dih.INDICE
    for c, valor in enumerate(hdr):
        if valor not in (None, ""):
            ws.write_string(0, c, str(valor), fmt_cab)
    for c, valor in enumerate(idx):
        if isinstance(valor, (int, float)):
            ws.write_number(1, c, valor)

    r = 2
    for fila in ws_jira.iter_rows(min_row=5, values_only=True):
        for c, valor in enumerate(fila):
            volcar(ws, r, c, valor, None, fmt_fecha)
        r += 1

    ws.autofilter(1, 0, max(r - 1, 2), 205)  # A2:GX...
    print(f"    DATA IN: {r - 2:,} filas de JIRA")


# ----------------------------------------------------------------------
# LEGAL VDF / OSP (desde Legal_Final)
# ----------------------------------------------------------------------
def escribir_legal(wb, nombre: str, ws_origen, etiqueta: str,
                   fmt_fecha, fmt_cab) -> None:
    ws = wb.add_worksheet(nombre)
    ws.set_column(0, 0, 18)
    ws.set_column(1, 40, 12)

    it = ws_origen.iter_rows(values_only=True)
    fila1 = next(it)
    for c, valor in enumerate(fila1):
        volcar(ws, 0, c, valor, None, fmt_fecha)

    fila2 = next(it)
    e_emf = fila2[1] if len(fila2) > 1 and fila2[1] else f"EMF {etiqueta}"
    e_amt = fila2[12] if len(fila2) > 12 and fila2[12] else f"AMT {etiqueta}"
    e_psc = fila2[23] if len(fila2) > 23 and fila2[23] else f"PSC {etiqueta}"
    ws.merge_range(1, 1, 1, 11, str(e_emf), fmt_cab)
    ws.merge_range(1, 12, 1, 22, str(e_amt), fmt_cab)
    ws.merge_range(1, 23, 1, 33, str(e_psc), fmt_cab)

    filas = 2
    for r, fila in enumerate(it, start=2):
        for c, valor in enumerate(fila):
            volcar(ws, r, c, valor, fmt_cab if r == 2 else None, fmt_fecha)
        filas = r + 1
    print(f"    {nombre}: {filas:,} filas")


# ----------------------------------------------------------------------
# drive (desde seguimiento.xlsx)
# ----------------------------------------------------------------------
DRIVE_COLS = [2, 29, 74, 129]  # B, AC, BV, DY (1-based)


def leer_drive(ws_origen):
    it = ws_origen.iter_rows(values_only=True)
    hdr = next(it)
    filas = []
    for fila in it:
        filas.append(tuple(fila[c - 1] if c - 1 < len(fila) else None
                           for c in DRIVE_COLS))
    return hdr, filas


def escribir_drive(wb, hdr, filas, fmt_fecha, fmt_cab) -> None:
    ws = wb.add_worksheet("drive")
    ws.set_column(0, 0, 14)
    ws.set_column(1, 2, 22)
    ws.set_column(3, 3, 10)

    for j, c in enumerate(DRIVE_COLS):
        if c - 1 < len(hdr) and hdr[c - 1] not in (None, ""):
            ws.write_string(0, j, str(hdr[c - 1]), fmt_cab)
    ws.write_string(1, 0, "", fmt_cab)
    ws.write_string(1, 1, "Fecha", fmt_cab)
    ws.write_string(1, 2, "Fecha", fmt_cab)
    ws.write_string(1, 3, "N/A", fmt_cab)

    for r, fila in enumerate(filas, start=2):
        for j, valor in enumerate(fila):
            volcar(ws, r, j, valor, None, fmt_fecha)
    print(f"    drive: {len(filas):,} filas")


# ----------------------------------------------------------------------
# Leg_Wn (desde Revision -> Detalle)
# ----------------------------------------------------------------------
LEG_COLS = list(range(1, 10)) + [13] + list(range(22, 27))  # A:I, M, V:Z


def escribir_leg_wn(wb, ws_origen, fmt_fecha, fmt_cab) -> None:
    ws = wb.add_worksheet("Leg_Wn")
    ws.set_column(0, 0, 10)
    ws.set_column(1, 8, 14)
    ws.set_column(9, 9, 22)
    ws.set_column(10, 13, 22)

    r = 0
    for fila in ws_origen.iter_rows(values_only=True):
        for j, c in enumerate(LEG_COLS):
            valor = fila[c - 1] if c - 1 < len(fila) else None
            volcar(ws, r, j, valor, fmt_cab if r == 0 else None, fmt_fecha)
        r += 1
    print(f"    Leg_Wn: {r:,} filas")


# ----------------------------------------------------------------------
# ZOOM
# ----------------------------------------------------------------------
def _criterios_base(campo: str, f_patron: int) -> str:
    return (f"'PREVISIONES JUAN'!${campo}:${campo},$B${f_patron},"
            f"'PREVISIONES JUAN'!$T:$T,\">0\","
            f"'PREVISIONES JUAN'!$N:$N,\"<>RIP\"")


def _f_semana(campo: str, f_patron: int, f_fechas: int, col: str) -> str:
    base = _criterios_base(campo, f_patron)
    rango = (f"'PREVISIONES JUAN'!$AV:$AV,\">=\"&{col}{f_fechas},"
             f"'PREVISIONES JUAN'!$AV:$AV,\"<=\"&({col}{f_fechas}+6)")
    partes = [f"COUNTIFS({base},{rango},'PREVISIONES JUAN'!$AU:$AU,\"{au}\")"
              for au in ("", "-")]
    return "=" + "+".join(partes)


def _f_mas5(campo: str, f_patron: int, f_fechas: int) -> str:
    base = _criterios_base(campo, f_patron)
    limite = f"F{f_fechas}+6"
    partes = [f"COUNTIFS({base},'PREVISIONES JUAN'!$AV:$AV,\">\"&{limite},"
              f"'PREVISIONES JUAN'!$AU:$AU,\"{au}\")" for au in ("", "-")]
    return "=" + "+".join(partes)


def _f_no_plan(campo: str, f_patron: int, f_inicio: int) -> str:
    """No planificada: sin AV, o AV anterior al lunes actual; y sin AU."""
    base = _criterios_base(campo, f_patron)
    partes = []
    for av in ("", "-"):
        for au in ("", "-"):
            partes.append(
                f"COUNTIFS({base},'PREVISIONES JUAN'!$AV:$AV,\"{av}\","
                f"'PREVISIONES JUAN'!$AU:$AU,\"{au}\")"
            )
    for au in ("", "-"):
        partes.append(
            f"COUNTIFS({base},'PREVISIONES JUAN'!$AV:$AV,\"<\"&B{f_inicio},"
            f"'PREVISIONES JUAN'!$AU:$AU,\"{au}\")"
        )
    return "=" + "+".join(partes)


def escribir_zoom(wb) -> None:
    ws = wb.add_worksheet("ZOOM")
    ws.set_column(0, 0, 34)
    ws.set_column(1, 5, 12)
    ws.set_column(6, 6, 16)
    ws.set_column(7, 7, 12)
    ws.set_column(8, 8, 10)

    f_titulo = wb.add_format({"bold": True, "font_size": 14, "font_color": "#FFFFFF",
                              "bg_color": "#1F4E78", "align": "center",
                              "valign": "vcenter"})
    f_sub = wb.add_format({"italic": True, "font_color": "#555555", "text_wrap": True})
    f_esc = wb.add_format({"bold": True, "font_size": 14, "bg_color": "#D9E1F2",
                           "border": 1, "align": "center", "valign": "vcenter"})
    f_lbl = wb.add_format({"bold": True, "border": 1})
    f_pat = wb.add_format({"bg_color": "#FFF2CC", "border": 1})
    f_cab = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1,
                           "align": "center", "valign": "vcenter", "text_wrap": True})
    f_fecha = wb.add_format({"num_format": "dd/mm/yyyy", "border": 1,
                             "align": "center", "bg_color": "#D9E1F2"})
    f_num = wb.add_format({"border": 1, "align": "center"})

    ws.merge_range(0, 0, 0, 8, "ZOOM · Incidencias planificadas (Previsión RFOA)", f_titulo)
    ws.merge_range(
        1, 0, 1, 8,
        "Semana 1 = lunes de la semana actual. Solo se cuentan incidencias con Fecha PO. "
        "No planificada = sin Previsión RFOA, sin Fecha RFOA, o Previsión RFOA anterior a esta semana. "
        "No se cuentan como planificadas las ya ejecutadas en semanas anteriores. "
        "Total = 5 semanas + No planificada + > 5 sem.",
        f_sub,
    )

    fila = 3
    for campo, patron, nombre, descripcion in ESCENARIOS:
        ws.merge_range(fila, 0, fila, 8, nombre, f_esc)

        fila_patron = fila + 1
        ws.set_row(fila_patron, None, None, {"hidden": True})
        ws.write_string(fila_patron, 0, "Patrón", f_lbl)
        ws.write_string(fila_patron, 1, patron, f_pat)
        ws.write_string(fila_patron, 3, "Campo", f_lbl)
        ws.write_string(fila_patron, 4, "Proyectos Radio" if campo == "G"
                        else "Proyecto & Subproyecto", f_lbl)

        # Fila oculta con las fechas reales (los COUNTIFS necesitan fechas)
        fila_fechas = fila + 2
        ws.set_row(fila_fechas, None, None, {"hidden": True})
        ws.write_formula(fila_fechas, 1, "=TODAY()-WEEKDAY(TODAY(),2)+1", f_fecha)
        for k in range(2, 6):
            prev = get_column_letter(k)
            ws.write_formula(fila_fechas, k, f"={prev}{fila_fechas + 1}+7", f_fecha)

        # Cabecera visible: WXX_AAAA
        fila_cab = fila + 3
        ws.write_string(fila_cab, 0, "Semana", f_cab)
        for k in range(1, 6):
            col = get_column_letter(k + 1)
            ws.write_formula(
                fila_cab, k,
                f'="W"&TEXT(WEEKNUM({col}{fila_fechas + 1},21),"00")&"_"'
                f'&YEAR({col}{fila_fechas + 1}-WEEKDAY({col}{fila_fechas + 1},2)+4)',
                f_cab,
            )
        ws.write_string(fila_cab, 6, "No planificada", f_cab)
        ws.write_string(fila_cab, 7, "> 5 sem", f_cab)
        ws.write_string(fila_cab, 8, "Total", f_cab)

        fila_datos = fila + 4
        ws.write_string(fila_datos, 0, "Nº de Tickets JIRA", f_lbl)
        for k in range(1, 6):
            col = get_column_letter(k + 1)
            ws.write_formula(
                fila_datos, k,
                _f_semana(campo, fila_patron + 1, fila_fechas + 1, col),
                f_num,
            )
        ws.write_formula(
            fila_datos, 6,
            _f_no_plan(campo, fila_patron + 1, fila_fechas + 1),
            f_num,
        )
        ws.write_formula(
            fila_datos, 7,
            _f_mas5(campo, fila_patron + 1, fila_fechas + 1),
            f_num,
        )
        ws.write_formula(
            fila_datos, 8,
            f"=SUM(B{fila_datos + 1}:F{fila_datos + 1})"
            f"+G{fila_datos + 1}+H{fila_datos + 1}",
            f_num,
        )
        fila += 6

    print("    ZOOM: 6 escenarios")


# ----------------------------------------------------------------------
# DETALLE_TICKETS (hoja oculta de drill-down)
# ----------------------------------------------------------------------
def _semana_txt(av, au, lunes: date) -> str:
    if isinstance(av, datetime):
        av = av.date()
    if isinstance(au, datetime):
        au = au.date()
    if not isinstance(av, date):
        return "Sin previsión"
    if av < lunes:
        if not isinstance(au, date):
            return "Sin previsión"
        return "Pasada"
    if av > lunes + timedelta(days=34):
        return "> 5 sem"
    iso = av.isocalendar()
    return f"W{iso[1]:02d}_{iso[0]}"


def escribir_detalle(wb, jmap: dict, fmt_fecha, fmt_cab) -> None:
    ws = wb.add_worksheet("DETALLE_TICKETS")
    titulos = ["Escenario", "IR", "Semana previsión", "Estado", "Fecha PO",
               "Previsión RFOA", "Fecha RFOA", "Proyectos Radio",
               "Proyecto & Subproyecto"]
    for c, t in enumerate(titulos):
        ws.write_string(0, c, t, fmt_cab)
    ws.set_column(0, 0, 42)
    ws.set_column(1, 1, 14)
    ws.set_column(2, 2, 16)
    ws.set_column(3, 3, 20)
    ws.set_column(4, 6, 14)
    ws.set_column(7, 8, 45)
    ws.freeze_panes(1, 0)

    lunes = date.today() - timedelta(days=date.today().weekday())
    r = 1
    for ir, d in jmap.items():
        t = d.get("t")
        if not isinstance(t, (date, datetime)):
            continue
        if isinstance(d.get("au"), (date, datetime)):
            continue
        if str(d.get("estado") or "").strip().upper() == "RIP":
            continue
        g = str(d.get("g") or "")
        e = str(d.get("e") or "")
        nombres = []
        for campo, patron, nombre, _ in ESCENARIOS:
            texto = g if campo == "G" else e
            if _coincide(patron, texto) and nombre not in nombres:
                nombres.append(nombre)
        datos = [" + ".join(nombres), ir, _semana_txt(d.get("av"), d.get("au"), lunes),
                 d.get("estado"), t, d.get("av"), d.get("au"), g, e]
        for c, v in enumerate(datos):
            volcar(ws, r, c, v, None, fmt_fecha)
        r += 1
    print(f"    DETALLE_TICKETS: {r - 1:,} filas")



# ----------------------------------------------------------------------
# Informe de cambios respecto al archivo anterior
# ----------------------------------------------------------------------
def escribir_informe(previas, irs, calc) -> None:
    destino = DESTINO.parent / "Informe_Cambios.xlsx"
    wb = xlsxwriter.Workbook(str(destino))
    f_cab = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
    f_txt = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})

    anterior = {}
    for fila in previas:
        ir = str(fila[0]).strip() if fila and fila[0] else ""
        if ir:
            anterior[ir] = (fila[7], fila[27], fila[35])

    cambios, actualizados, nuevos = [], [], []
    for ir in irs:
        h, ab, aj = calc.get(ir, ("", "Sin IR", "N/A"))
        if ir not in anterior:
            nuevos.append([ir, h, ab, aj])
            continue
        ah, aab, aaj = anterior[ir]
        if str(ah or "") != str(h or ""):
            cambios.append([ir, ah, h])
        none = (None, "", "Sin IR")
        if (aab in none and ab not in none) or (aaj in none and aj not in none):
            actualizados.append([ir, aab, ab, aaj, aj])

    def hoja(nombre, cabeceras, filas):
        ws = wb.add_worksheet(nombre)
        for c, t in enumerate(cabeceras):
            ws.write_string(0, c, t, f_cab)
        for r, fila in enumerate(filas, start=1):
            for c, v in enumerate(fila):
                ws.write_string(r, c, "" if v is None else str(v), f_txt)
        for c in range(len(cabeceras)):
            ws.set_column(c, c, 28)
        return ws

    hoja("ESCENARIO_CAMBIADO", ["IR", "H anterior", "H actual"], cambios)
    hoja("IR_TX_AJ_ACTUALIZADO",
         ["IR", "AB anterior", "AB actual", "AJ anterior", "AJ actual"], actualizados)
    hoja("NUEVOS", ["IR", "H", "AB", "AJ"], nuevos)
    wb.close()
    print(f"    Informe_Cambios.xlsx: {len(cambios)} cambios H, "
          f"{len(actualizados)} AB/AJ, {len(nuevos)} nuevos")




# ----------------------------------------------------------------------
# Recálculo con Excel (Windows) y conversión a valores
# ----------------------------------------------------------------------
def recalcular_con_excel(path: Path) -> bool:
    """Abre el libro en Excel invisible, recalcula y guarda. Solo Windows."""
    if os.name != "nt":
        return False
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        return False

    excel = None
    try:
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        libro = excel.Workbooks.Open(str(path))
        excel.CalculateFullRebuild()
        libro.Save()
        libro.Close(SaveChanges=True)
        return True
    except Exception as e:  # noqa: BLE001
        print(f"    [AVISO] No se pudo recalcular con Excel: {e}")
        return False
    finally:
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def convertir_a_valores(path: Path) -> None:
    """Sustituye fórmulas por sus valores cacheados tras el recálculo."""
    wb = openpyxl.load_workbook(path)
    wv = openpyxl.load_workbook(path, data_only=True)
    for nombre in wb.sheetnames:
        ws, wsv = wb[nombre], wv[nombre]
        for fila in ws.iter_rows(min_row=2):
            for celda in fila:
                v = celda.value
                if isinstance(v, ArrayFormula) or (isinstance(v, str) and v.startswith("=")):
                    celda.value = wsv[celda.coordinate].value
    wb.save(path)
    wb.close()
    wv.close()


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main() -> None:
    t0 = time.time()
    print("=" * 70)
    print("  FASE 5 · Generando Seguimiento_Previsiones_Limpio.xlsx")
    print("=" * 70)

    requeridos = [JIRA, LEGAL, SEGUIMIENTO, REVISION, HOJA]
    faltan = [p for p in requeridos if not p.is_file()]
    if faltan:
        for p in faltan:
            print("  Falta:", p)
        sys.exit("No se puede generar: faltan archivos de entrada.")

    print("[1] Leyendo fuentes...")
    jmap = leer_jira_map()
    print(f"    JIRA: {len(jmap):,} IRs")
    headers, previas = leer_anterior(ANTERIOR)
    if previas:
        print(f"    Anterior: {len(previas):,} filas de {ANTERIOR.name}")
    else:
        print("    [AVISO] Sin archivo anterior: se genera desde cero, SIN heredar fórmulas.")
        headers = list(ph.HEADERS)
    tx, act = leer_subtareas()
    print(f"    Subtareas: Tx Single RAN={len(tx):,} | ACTIVACION OR={len(act):,}")

    wb_jira = openpyxl.load_workbook(JIRA, read_only=True, data_only=True)
    wb_legal = openpyxl.load_workbook(LEGAL, read_only=True, data_only=True)
    wb_seg = openpyxl.load_workbook(SEGUIMIENTO, read_only=True, data_only=True)
    wb_rev = openpyxl.load_workbook(REVISION, read_only=True, data_only=True)

    drive_hdr, drive_rows = leer_drive(wb_seg["NUEVO SEGUIMIENTO"])
    irs = construir_irs(previas, drive_rows)
    print(f"    IRs: {len(previas):,} anteriores + "
          f"{len(irs) - len(previas):,} nuevos = {len(irs):,}")
    calc = {ir: calcular_manual(ir, jmap, tx, act) for ir in irs}

    print("[2] Escribiendo...")
    wb = xlsxwriter.Workbook(str(DESTINO), {"constant_memory": True,
                                            "strings_to_numbers": False})
    fmt_fecha = wb.add_format({"num_format": "dd/mm/yyyy"})
    fmt_cab = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1,
                             "align": "center", "valign": "vcenter"})

    escribir_data_in(wb, wb_jira["IR_VDF"], fmt_fecha, fmt_cab)
    escribir_legal(wb, "LEGAL OSP", wb_legal["legal_osp"], "OSP", fmt_fecha, fmt_cab)
    escribir_legal(wb, "LEGAL VDF", wb_legal["legal_vdf"], "VDF", fmt_fecha, fmt_cab)
    escribir_previsiones(wb, headers, previas, irs, calc)
    escribir_informe(previas, irs, calc)
    escribir_zoom(wb)
    escribir_detalle(wb, jmap, fmt_fecha, fmt_cab)
    escribir_drive(wb, drive_hdr, drive_rows, fmt_fecha, fmt_cab)
    escribir_leg_wn(wb, wb_rev["Detalle"], fmt_fecha, fmt_cab)

    wb.close()
    for w in (wb_jira, wb_legal, wb_seg, wb_rev):
        w.close()

    if "--sin-recalcular" not in sys.argv:
        if recalcular_con_excel(DESTINO):
            convertir_a_valores(DESTINO)
            print("    Recalculado con Excel y convertido a valores")
        else:
            print("    Sin Excel: calculando valores con pycel por bloques...")
            try:
                from calc_pycel import calcular_valores
                vals = calcular_valores(
                    DESTINO,
                    progreso=lambda hechas, total: print(
                        f"      pycel: {hechas}/{total} filas", end="\r"),
                )
                print()
                wb_val = openpyxl.load_workbook(DESTINO)
                ws_val = wb_val["PREVISIONES JUAN"]
                for (fila, col), valor in vals.items():
                    ws_val.cell(row=fila, column=col, value=valor)
                wb_val.save(DESTINO)
                wb_val.close()
                print("    Calculado con pycel y convertido a valores")
            except Exception as e:  # noqa: BLE001
                print(f"    [AVISO] No se pudo calcular con pycel: {e}")
                print("    Se conservan las fórmulas; instala pycel: "
                      "py -m pip install pycel")

    tam = DESTINO.stat().st_size / 1e6
    print("\n" + "=" * 70)
    print(f"  [OK] {DESTINO.name} ({tam:.1f} MB) en {time.time() - t0:.1f}s")
    print("       DATA IN · LEGAL OSP · LEGAL VDF · PREVISIONES JUAN · ZOOM · "
          "DETALLE_TICKETS · drive · Leg_Wn")
    print("=" * 70)


if __name__ == "__main__":
    main()
