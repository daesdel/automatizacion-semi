#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FASE 8 - Adicionales (libro limpio).

Genera ADICIONALES_SEMI_Limpio.xlsx con solo:
  - DATA IN: datos frescos de JIRA.xlsx desde A3 (cabeceras fila 1,
    índice fila 2).
  - PLANIFICACION: 4 rangos + resumen AG2:AM16, sin gráficos ni ocultas.

Lógica:
  1) Foto fija WN-1 si no existe (copia del original, solo lectura).
  2) Rango 2 (I:O) <- Rango 3 (Q:W) de la foto fija.
  3) Rango 1 (B:G) <- DATA IN con Prevision RFOA de la semana N-1,
     descontando los IR del Rango 2.
  4) Rango 3 (Q:W) <- DATA IN con Prevision RFOA de la semana N.
  5) Rango 4 (Y:AD) <- DATA IN con Prevision RFOA de la semana N+1.

Multiplataforma: xlsxwriter. No usa Excel COM ni modifica el original.
"""

import shutil
import sys
import warnings
from datetime import date, datetime, timedelta
from pathlib import Path

import xlsxwriter
from python_calamine import CalamineWorkbook
from xlsxwriter.utility import xl_col_to_name

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

HERE = Path(__file__).resolve().parent
BASE = HERE.parent / "Nueva carpeta (2)"


def _arg(nombre, defecto):
    pref = f"--{nombre}="
    for a in sys.argv[1:]:
        if a.startswith(pref):
            return Path(a.split("=", 1)[1]).expanduser()
    return defecto


ADICIONALES = _arg("adicionales", BASE / "ADICIONALES SEGUIMIENTO.xlsx")
BACKUP = _arg("backup", BASE / "ADICIONALES SEGUIMIENTO_WN-1.xlsx")
JIRA = _arg("jira", BASE / "JIRA.xlsx")
OUT = _arg("out", BASE / "resultados" / "ADICIONALES_SEMI_Limpio.xlsx")

HOJA_PLAN = "PLANIFICACION"
FILA_DATOS = 3

PROYECTOS = [
    "Plan Capacidad - Upgrade",
    "Traslado",
    "JUMPING - Capacidad radio (Portadoras)",
    "Plan Capacidad - EVENTOS",
    "Nueva Cobertura - Sin Subproyecto",
    "Cliente",
    "Forced decommissioning",
    "Plan 5G - Commitment",
    "Repliegue",
    "JUMPING - Jumping Radio",
    "JUMPING - Regularizacion Baseline",
    "Plan 5G - NR700 EA",
]


def _fecha(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _lunes_actual():
    hoy = date.today()
    return hoy - timedelta(days=hoy.weekday())


def _en_semana(v, offset):
    d = _fecha(v)
    if d is None:
        return False
    ini = _lunes_actual() + timedelta(days=7 * offset)
    return ini <= d <= ini + timedelta(days=6)


def leer_jira():
    wb = CalamineWorkbook.from_path(str(JIRA))
    filas = wb.get_sheet_by_name("IR_VDF").to_python(skip_empty_area=True)
    hdr = filas[3]
    datos = filas[4:]
    idx = {h: i for i, h in enumerate(hdr) if h}

    def col(nombre, fila):
        i = idx.get(nombre)
        return fila[i] if i is not None and i < len(fila) else None

    registros = []
    for fila in datos:
        if col("Clave", fila) in (None, ""):
            continue
        registros.append({
            "ir": str(col("Clave", fila)).strip(),
            "sr": col("Codigo SR", fila),
            "proy": col("Proyecto & Subproyecto", fila),
            "radio": col("Proyectos Radio", fila),
            "prev": col("Prevision RFOA", fila),
            "real": col("Real RFOA", fila),
        })
    return hdr[:261], [list(f[:261]) for f in datos], registros


def leer_rango2(backup: Path):
    wb = CalamineWorkbook.from_path(str(backup))
    filas = wb.get_sheet_by_name(HOJA_PLAN).to_python(skip_empty_area=True)
    salida = []
    for fila in filas[2:]:
        ir = fila[16] if len(fila) > 16 else None  # Q
        if ir in (None, ""):
            continue
        salida.append([
            ir,
            fila[17] if len(fila) > 17 else None,
            fila[18] if len(fila) > 18 else None,
            "RFOA",
            fila[20] if len(fila) > 20 else None,
            fila[21] if len(fila) > 21 else None,
            fila[22] if len(fila) > 22 else None,
        ])
    return salida


def construir(regs, rango2):
    claves2 = {str(r[0]).strip() for r in rango2}
    r1, r3, r4 = [], [], []
    for d in regs:
        if d["ir"] not in claves2 and _en_semana(d["prev"], -1):
            r1.append([d["ir"], d["sr"], d["proy"], "RFOA", d["radio"], d["real"]])
        if _en_semana(d["prev"], 0):
            r3.append([d["ir"], d["sr"], d["proy"], "RFOA", d["radio"],
                       d["prev"], d["real"]])
        if _en_semana(d["prev"], 1):
            r4.append([d["ir"], d["sr"], d["proy"], "RFOA", d["radio"], d["prev"]])
    return r1, r3, r4


def _val(v):
    return None if v == "" else v


def _escribir_bloque(ws, fmt_fecha, celda, filas, ncols):
    from openpyxl.utils import column_index_from_string
    c0 = column_index_from_string(celda[0])
    r0 = int(celda[1:])
    for i, fila in enumerate(filas):
        for j in range(ncols):
            v = _val(fila[j]) if j < len(fila) else None
            if isinstance(v, (datetime, date)):
                ws.write_datetime(r0 - 1 + i, c0 - 1 + j, v, fmt_fecha)
            else:
                ws.write(r0 - 1 + i, c0 - 1 + j, v)


def construir_libro(hdr, filas_data, r1, r2, r3, r4):
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(OUT), {"constant_memory": True,
                                        "strings_to_numbers": False})
    f_cab = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1,
                           "align": "center", "valign": "vcenter", "text_wrap": True})
    f_tit = wb.add_format({"bold": True, "bg_color": "#1F4E78",
                           "font_color": "#FFFFFF", "border": 1,
                           "align": "center", "valign": "vcenter"})
    f_fecha = wb.add_format({"num_format": "dd/mm/yyyy"})
    f_num = wb.add_format({"border": 1, "align": "center"})
    f_txt = wb.add_format({"border": 1})

    # DATA IN
    wd = wb.add_worksheet("DATA IN")
    for j, h in enumerate(hdr, 1):
        wd.write_string(0, j - 1, str(h) if h is not None else "", f_cab)
        wd.write_number(1, j - 1, j, f_num)
    for i, fila in enumerate(filas_data, start=2):
        for j, v in enumerate(fila, 1):
            if isinstance(v, (datetime, date)):
                wd.write_datetime(i, j - 1, v, f_fecha)
            else:
                wd.write(i, j - 1, v)
    wd.freeze_panes(2, 1)
    wd.set_column(0, 0, 16)

    # PLANIFICACION
    wp = wb.add_worksheet(HOJA_PLAN)
    wp.set_column(0, 0, 3)
    wp.set_column(1, 5, 16)
    wp.set_column(6, 7, 3)
    wp.set_column(8, 13, 16)
    wp.set_column(14, 15, 3)
    wp.set_column(16, 21, 16)
    wp.set_column(22, 23, 3)
    wp.set_column(24, 28, 16)
    wp.set_column(29, 31, 3)
    wp.set_column(32, 38, 14)

    # Fila 1: títulos de bloque
    for rango, titulo in (("B1:G1", "SEMANA N-1 (no plan)"),
                          ("I1:O1", "SEMANA N-1"),
                          ("Q1:W1", "SEMANA N"),
                          ("Y1:AD1", "SEMANA N+1")):
        wp.merge_range(rango, titulo, f_tit)

    # Fila 2: cabeceras de los bloques y del resumen
    from openpyxl.utils import column_index_from_string
    cabeceras = [
        (1, ["JIRA IR", "SR", "PROYECTO & SUBPROYECTO", "PIVOTE",
             "Proyectos Radio", "Real RFOA"]),
        (8, ["JIRA IR", "SR", "PROYECTO & SUBPROYECTO", "PIVOTE",
             "Proyectos Radio", "Prev RFOA JIRA", "Real RFOA"]),
        (16, ["JIRA IR", "SR", "PROYECTO & SUBPROYECTO", "PIVOTE",
              "Proyectos Radio", "Prev RFOA JIRA", "Real RFOA"]),
        (24, ["JIRA IR", "SR", "PROYECTO & SUBPROYECTO", "PIVOTE",
              "Proyectos Radio", "Prev RFOA JIRA"]),
    ]
    for c0, nombres in cabeceras:
        for j, nombre in enumerate(nombres):
            wp.write_string(1, c0 + j, nombre, f_cab)
    wp.write_string(1, 32, "PROYECTO", f_tit)
    wp.merge_range("AH2:AJ2", "W n-1", f_tit)
    wp.merge_range("AK2:AL2", "Wn", f_tit)
    wp.write_string("AM2", "Wn+1", f_tit)

    max_filas = max(len(r1), len(r2), len(r3), len(r4), 1)
    filas_a_escribir = max(max_filas, 14)  # hasta la fila 16 (resumen)

    for i in range(filas_a_escribir):
        fila = 2 + i  # 0-based: fila Excel = i+3
        for c0, datos, ncols in ((1, r1, 6), (8, r2, 7),
                                 (16, r3, 7), (24, r4, 6)):
            if i >= len(datos):
                continue
            for j in range(ncols):
                v = _val(datos[i][j]) if j < len(datos[i]) else None
                if isinstance(v, (datetime, date)):
                    wp.write_datetime(fila, c0 + j, v, f_fecha)
                else:
                    wp.write(fila, c0 + j, v)

        excel_row = fila + 1
        if excel_row == 3:
            for celda, texto in (("AH3", "PLANIFICADO Wn-1"),
                                 ("AI3", "REALIZADO Wn-1"),
                                 ("AJ3", "REALIZADO NO PLAN Wn-1"),
                                 ("AK3", "PLANIFICADO Wn"),
                                 ("AL3", "REALIZADO Wn"),
                                 ("AM3", "PLANIFICADO Wn+1")):
                wp.write_string(celda, texto, f_cab)
        elif 4 <= excel_row <= 15:
            proyecto = PROYECTOS[excel_row - 4]
            wp.write_string(fila, 32, proyecto, f_txt)
            wp.write_formula(fila, 33,
                             f'=IF($AG{excel_row}<>"",COUNTIFS($K:$K,"*"&$AG{excel_row}&"*",$L:$L,"*RFOA*"),"")', f_num)
            wp.write_formula(fila, 34,
                             f'=IF($AG{excel_row}<>"",COUNTIFS($K:$K,"*"&$AG{excel_row}&"*",$L:$L,"*RFOA*",$O:$O,">0"),"")', f_num)
            wp.write_formula(fila, 35,
                             f'=IF($AG{excel_row}<>"",COUNTIFS($D:$D,"*"&$AG{excel_row}&"*",$E:$E,"*RFOA*",$G:$G,">0"),"")', f_num)
            wp.write_formula(fila, 36,
                             f'=IF($AG{excel_row}<>"",COUNTIF($S:$S,"*"&$AG{excel_row}&"*"),"")', f_num)
            wp.write_formula(fila, 37,
                             f'=IF($AG{excel_row}<>"",COUNTIFS($S:$S,"*"&$AG{excel_row}&"*",$W:$W,">0"),"")', f_num)
            wp.write_formula(fila, 38,
                             f'=IF($AG{excel_row}<>"",COUNTIF($AA:$AA,"*"&$AG{excel_row}&"*"),"")', f_num)
        elif excel_row == 16:
            wp.write_string(fila, 32, "TOTAL", f_cab)
            for col in range(33, 39):
                letra = xl_col_to_name(col)
                wp.write_formula(fila, col, f"=SUM({letra}4:{letra}15)", f_num)

    wp.freeze_panes(2, 1)

    wb.close()
    return OUT


def main():
    print("=" * 70)
    print("  FASE 8 · Adicionales (libro limpio)")
    print("=" * 70)
    if not JIRA.is_file():
        sys.exit(f"No existe {JIRA}")
    fuente = BACKUP if BACKUP.is_file() else ADICIONALES
    if not fuente.is_file():
        sys.exit(f"No existe {ADICIONALES} ni {BACKUP}")

    hdr, filas_data, regs = leer_jira()
    rango2 = leer_rango2(fuente)
    r1, r3, r4 = construir(regs, rango2)
    print(f"  JIRA: {len(regs):,} tickets")
    print(f"  Rangos: R1={len(r1)}, R2={len(rango2)}, R3={len(r3)}, R4={len(r4)}")

    if "--dry-run" in sys.argv:
        print("  [DRY-RUN] No se escribe ningún archivo.")
        return

    if not BACKUP.is_file():
        shutil.copy2(ADICIONALES, BACKUP)
        print(f"  Foto fija creada: {BACKUP.name} (original intacto)")

    construir_libro(hdr, filas_data, r1, rango2, r3, r4)
    print(f"  [OK] {OUT} ({OUT.stat().st_size / 1e6:.1f} MB)")
    print("       Hojas: DATA IN · PLANIFICACION")


if __name__ == "__main__":
    main()
