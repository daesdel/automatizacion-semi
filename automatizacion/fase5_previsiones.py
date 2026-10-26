#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FASE 5 - Archivo limpio: Seguimiento_Previsiones_Limpio.xlsx
=============================================================

Genera el libro con 7 hojas a partir de:

    plantilla_previsiones.xlsx   -> PREVISIONES JUAN (fórmulas + formato) y
                                    cabeceras/índice de DATA IN
    JIRA.xlsx (Fase 1)           -> datos de DATA IN
    resultados/Legal_Final.xlsx  -> LEGAL VDF / LEGAL OSP
    resultados/seguimiento.xlsx  -> drive
    Revision Legalizacion ...    -> Leg_Wn
    (ZOOM se calcula)

Ya no se lee el SEGUIMIENTO PREVISIONES.xlsx original.

Uso:
    python automatizacion/fase5_previsiones.py
    python automatizacion/fase5_previsiones.py --out=RUTA
"""

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

HERE = Path(__file__).resolve().parent
BASE_DIR = HERE.parent / "Nueva carpeta (2)"


def _arg_path(nombre: str, defecto: Path) -> Path:
    pref = f"--{nombre}="
    for a in sys.argv[1:]:
        if a.startswith(pref):
            return Path(a.split("=", 1)[1]).expanduser()
    return defecto


PLANTILLA = _arg_path("plantilla", BASE_DIR / "plantilla_previsiones.xlsx")
JIRA = _arg_path("jira", BASE_DIR / "JIRA.xlsx")
LEGAL = _arg_path("legal", BASE_DIR / "resultados" /
                  f"{date.today():%d%m%Y}_legal ministerio.xlsx")
SEGUIMIENTO = _arg_path("seguimiento", BASE_DIR / "resultados" / "seguimiento.xlsx")
REVISION = _arg_path("revision", BASE_DIR / "Revision Legalizacion SEMI SemN.xlsx")
DESTINO = _arg_path("out", BASE_DIR / "resultados" / "Seguimiento_Previsiones_Limpio.xlsx")

sys.path.insert(0, str(HERE))
import estilos_previsiones as est  # noqa: E402
import data_in_headers as dih  # noqa: E402

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
# PREVISIONES JUAN (desde la plantilla)
# ----------------------------------------------------------------------
def escribir_previsiones(wb, ws_origen) -> None:
    ws = wb.add_worksheet("PREVISIONES JUAN")
    for col, ancho in est.ANCHOS.items():
        ws.set_column(col - 1, col - 1, ancho)
    ws.set_row(0, est.ALTO_FILA1)
    ws.freeze_panes(*est.FREEZE)

    cab, dat = est.crear_formatos(wb)
    filas = 0
    with FormulasNormales():
        for r, fila in enumerate(ws_origen.iter_rows(values_only=True)):
            for c, valor in enumerate(fila, start=1):
                fmt = cab.get(c) if r == 0 else dat.get(c)
                volcar(ws, r, c - 1, valor, fmt)
            filas = r + 1

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
    print(f"    PREVISIONES JUAN: {filas:,} filas (plantilla)")


# ----------------------------------------------------------------------
# DATA IN: cabeceras de plantilla + datos de JIRA
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


def escribir_drive(wb, ws_origen, fmt_fecha, fmt_cab) -> None:
    ws = wb.add_worksheet("drive")
    ws.set_column(0, 0, 14)
    ws.set_column(1, 2, 22)
    ws.set_column(3, 3, 10)

    it = ws_origen.iter_rows(values_only=True)
    hdr = next(it)
    for j, c in enumerate(DRIVE_COLS):
        if c - 1 < len(hdr) and hdr[c - 1] not in (None, ""):
            ws.write_string(0, j, str(hdr[c - 1]), fmt_cab)
    ws.write_string(1, 0, "", fmt_cab)
    ws.write_string(1, 1, "Fecha", fmt_cab)
    ws.write_string(1, 2, "Fecha", fmt_cab)
    ws.write_string(1, 3, "N/A", fmt_cab)

    r = 2
    for fila in it:
        for j, c in enumerate(DRIVE_COLS):
            valor = fila[c - 1] if c - 1 < len(fila) else None
            volcar(ws, r, j, valor, None, fmt_fecha)
        r += 1
    print(f"    drive: {r - 2:,} filas")


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
            f"'PREVISIONES JUAN'!$T:$T,\">0\"")


def _f_semana(campo: str, f_patron: int, f_fechas: int, col: str) -> str:
    base = _criterios_base(campo, f_patron)
    rango = (f"'PREVISIONES JUAN'!$AV:$AV,\">=\"&{col}{f_fechas},"
             f"'PREVISIONES JUAN'!$AV:$AV,\"<=\"&({col}{f_fechas}+6)")
    total = f"COUNTIFS({base},{rango})"
    hechas = (f"COUNTIFS({base},{rango},"
              f"'PREVISIONES JUAN'!$AU:$AU,\"<\"&B{f_fechas})")
    return f"={total}-{hechas}"


def _f_mas5(campo: str, f_patron: int, f_fechas: int) -> str:
    base = _criterios_base(campo, f_patron)
    limite = f"F{f_fechas}+6"
    total = f"COUNTIFS({base},'PREVISIONES JUAN'!$AV:$AV,\">\"&{limite})"
    hechas = (f"COUNTIFS({base},'PREVISIONES JUAN'!$AV:$AV,\">\"&{limite},"
              f"'PREVISIONES JUAN'!$AU:$AU,\"<\"&B{f_fechas})")
    return f"={total}-{hechas}"


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
def _semana_txt(av, lunes: date) -> str:
    if isinstance(av, datetime):
        av = av.date()
    if not isinstance(av, date):
        return "Sin previsión"
    if av < lunes:
        return "Pasada"
    if av > lunes + timedelta(days=34):
        return "> 5 sem"
    iso = av.isocalendar()
    return f"W{iso[1]:02d}_{iso[0]}"


def escribir_detalle(wb, ws_jira, fmt_fecha, fmt_cab) -> None:
    ws = wb.add_worksheet("DETALLE_TICKETS")
    titulos = ["Escenario", "IR", "Semana previsión", "Estado", "Fecha PO",
               "Previsión RFOA", "Fecha RFOA", "Proyectos Radio",
               "Proyecto & Subproyecto"]
    for c, t in enumerate(titulos):
        ws.write_string(0, c, t, fmt_cab)
    ws.set_column(0, 0, 22)
    ws.set_column(1, 1, 14)
    ws.set_column(2, 2, 14)
    ws.set_column(3, 3, 20)
    ws.set_column(4, 6, 14)
    ws.set_column(7, 8, 45)

    hdr = list(next(ws_jira.iter_rows(min_row=4, max_row=4, values_only=True)))
    idx = {h: i for i, h in enumerate(hdr) if h}
    lunes = date.today() - timedelta(days=date.today().weekday())

    def val(fila, nombre):
        i = idx.get(nombre)
        return fila[i] if i is not None and i < len(fila) else None

    r = 1
    for fila in ws_jira.iter_rows(min_row=5, values_only=True):
        t = val(fila, "Fecha PO")
        if not isinstance(t, (date, datetime)):
            continue
        g = val(fila, "Proyectos Radio")
        e = val(fila, "Proyecto & Subproyecto")
        for campo, patron, nombre, _ in ESCENARIOS:
            texto = g if campo == "G" else e
            if not _coincide(patron, texto):
                continue
            av = val(fila, "Prevision RFOA")
            datos = [nombre, val(fila, "Clave"), _semana_txt(av, lunes),
                     val(fila, "Estado"), t, av, val(fila, "Real RFOA"),
                     g, e]
            for c, v in enumerate(datos):
                if c in (2, 4, 5, 6):
                    volcar(ws, r, c, v, None, fmt_fecha)
                else:
                    volcar(ws, r, c, v, None, fmt_fecha)
            r += 1
    ws.hide()
    print(f"    DETALLE_TICKETS: {r - 1:,} filas (oculta)")



# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main() -> None:
    t0 = time.time()
    print("=" * 70)
    print("  FASE 5 · Generando Seguimiento_Previsiones_Limpio.xlsx")
    print("=" * 70)

    faltan = [p for p in (PLANTILLA, JIRA, LEGAL, SEGUIMIENTO, REVISION) if not p.is_file()]
    if faltan:
        for p in faltan:
            print("  Falta:", p)
        sys.exit("No se puede generar: faltan archivos de entrada.")

    print("[1] Abriendo plantilla y fuentes...")
    wb_plantilla = openpyxl.load_workbook(PLANTILLA, read_only=True, data_only=False)
    wb_jira = openpyxl.load_workbook(JIRA, read_only=True, data_only=True)
    wb_detalle = openpyxl.load_workbook(JIRA, read_only=True, data_only=True)
    wb_legal = openpyxl.load_workbook(LEGAL, read_only=True, data_only=True)
    wb_seg = openpyxl.load_workbook(SEGUIMIENTO, read_only=True, data_only=True)
    wb_rev = openpyxl.load_workbook(REVISION, read_only=True, data_only=True)

    print("[2] Escribiendo...")
    wb = xlsxwriter.Workbook(str(DESTINO), {"constant_memory": True,
                                            "strings_to_numbers": False})
    fmt_fecha = wb.add_format({"num_format": "dd/mm/yyyy"})
    fmt_cab = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1,
                             "align": "center", "valign": "vcenter"})

    escribir_data_in(wb, wb_jira["IR_VDF"], fmt_fecha, fmt_cab)
    escribir_legal(wb, "LEGAL OSP", wb_legal["legal_osp"], "OSP", fmt_fecha, fmt_cab)
    escribir_legal(wb, "LEGAL VDF", wb_legal["legal_vdf"], "VDF", fmt_fecha, fmt_cab)
    escribir_previsiones(wb, wb_plantilla["PREVISIONES JUAN"])
    escribir_zoom(wb)
    escribir_detalle(wb, wb_detalle["IR_VDF"], fmt_fecha, fmt_cab)
    escribir_drive(wb, wb_seg["NUEVO SEGUIMIENTO"], fmt_fecha, fmt_cab)
    escribir_leg_wn(wb, wb_rev["Detalle"], fmt_fecha, fmt_cab)

    wb.close()
    for w in (wb_plantilla, wb_jira, wb_detalle, wb_legal, wb_seg, wb_rev):
        w.close()

    tam = DESTINO.stat().st_size / 1e6
    print("\n" + "=" * 70)
    print(f"  [OK] {DESTINO.name} ({tam:.1f} MB) en {time.time() - t0:.1f}s")
    print("       DATA IN · LEGAL OSP · LEGAL VDF · PREVISIONES JUAN · ZOOM · drive · Leg_Wn")
    print("=" * 70)


if __name__ == "__main__":
    main()