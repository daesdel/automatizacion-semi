#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FASE 8 - Adicionales.

Pasos:
  0) Foto fija ADICIONALES SEGUIMIENTO_WN-1.xlsx (si no existe).
  1) DATA IN <- JIRA.xlsx (IR_VDF), desde A3.
  2) Rango 2 (I:N) <- Rango 3 (Q:W) de la foto fija.
  3) Rango 1 (B:G) <- DATA IN con Prevision RFOA de la semana N-1, sin duplicar Rango 2.
  4) Rango 3 (Q:W) <- DATA IN con Prevision RFOA de la semana N.
  5) Rango 4 (Y:AD) <- DATA IN con Prevision RFOA de la semana N+1.
  6) Guarda ADICIONALES SEGUIMIENTO.xlsx.

En Windows usa Excel COM (rápido y conserva gráficos). En Mac usa openpyxl
(más lento y puede perder gráficos): úsese sobre una copia para pruebas.
"""

import os
import shutil
import sys
from datetime import date, datetime, time, timedelta
from pathlib import Path

import openpyxl

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

HOJA_DATOS = "DATA IN"
HOJA_PLAN = "PLANIFICACION"
RANGO2_FILAS = 22
FILA_DATOS = 3
FILA_FIN = 24


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
    """Devuelve (filas_para_DATA_IN, registros)."""
    wb = openpyxl.load_workbook(JIRA, read_only=True, data_only=True)
    ws = wb["IR_VDF"]
    hdr = list(next(ws.iter_rows(min_row=4, max_row=4, values_only=True)))
    idx = {h: i for i, h in enumerate(hdr) if h}

    def col(nombre, fila):
        i = idx.get(nombre)
        return fila[i] if i is not None and i < len(fila) else None

    filas, regs = [], []
    for fila in ws.iter_rows(min_row=5, values_only=True):
        if col("Clave", fila) in (None, ""):
            continue
        filas.append(list(fila[:261]))
        regs.append({
            "ir": str(col("Clave", fila)).strip(),
            "sr": col("Codigo SR", fila),
            "proy": col("Proyecto & Subproyecto", fila),
            "radio": col("Proyectos Radio", fila),
            "prev": col("Prevision RFOA", fila),
            "real": col("Real RFOA", fila),
        })
    wb.close()
    return filas, regs


def leer_rango2(backup: Path):
    """Copia Q:V del Rango 3 de la foto fija (6 columnas)."""
    wb = openpyxl.load_workbook(backup, read_only=True, data_only=True)
    ws = wb[HOJA_PLAN]
    salida = []
    for r in range(FILA_DATOS, ws.max_row + 1):
        ir = ws.cell(row=r, column=17).value  # Q
        if ir in (None, ""):
            continue
        salida.append([
            ir,
            ws.cell(row=r, column=18).value,  # R
            ws.cell(row=r, column=19).value,  # S
            "RFOA",
            ws.cell(row=r, column=21).value,  # U
            ws.cell(row=r, column=22).value,  # V Prevision RFOA
        ])
    wb.close()
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


def _normalizar(v):
    return v


def _matriz(filas, ncols):
    return [[_normalizar(f[i]) if i < len(f) else "" for i in range(ncols)]
            for f in filas]


def actualizar_excel(r1, r3, r4, rango2, jira_filas):
    import pythoncom
    import win32com.client as win32

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    libro = None
    try:
        libro = excel.Workbooks.Open(str(ADICIONALES))
        wd = libro.Worksheets(HOJA_DATOS)
        wj = excel.Workbooks.Open(str(JIRA)).Worksheets("IR_VDF")

        # DATA IN desde A3
        ult = wj.Cells(wj.Rows.Count, 1).End(-4162).Row
        wd.Range(f"A3:IS{max(ult + 2, wd.UsedRange.Rows.Count)}").ClearContents()
        if ult >= 5:
            wd.Range(f"A3:IS{3 + ult - 5}").Value = wj.Range(f"A5:IS{ult}").Value

        maximo = max(len(r1), len(r3), len(r4), len(rango2), RANGO2_FILAS)
        extra = max(0, maximo - RANGO2_FILAS)
        if extra:
            libro.Worksheets(HOJA_PLAN).Rows(f"25:{24 + extra}").Insert()
        fin = FILA_FIN + extra
        wp = libro.Worksheets(HOJA_PLAN)

        def limpiar(desde, hasta):
            for col in desde:
                wp.Range(f"{col}{FILA_DATOS}:{col}{fin}").ClearContents()

        limpiar(["B", "C", "D", "E", "F", "G"])
        limpiar(["I", "J", "K", "L", "M", "N"])
        limpiar(["Q", "R", "S", "T", "U", "V", "W"])
        limpiar(["Y", "Z", "AA", "AB", "AC", "AD"])

        bloques = [
            ("B3", r1, 6), ("I3", rango2, 6), ("Q3", r3, 7), ("Y3", r4, 6),
        ]
        for celda, filas, ncols in bloques:
            if not filas:
                continue
            mat = _matriz(filas, ncols)
            ini = int(celda[1:])
            col0 = celda[0]
            col1 = chr(ord(col0) + ncols - 1) if ncols <= 26 else col0
            # Para bloques de una sola letra (B, I, Q, Y)
            from openpyxl.utils import get_column_letter, column_index_from_string
            ci = column_index_from_string(col0)
            col1 = get_column_letter(ci + ncols - 1)
            wp.Range(f"{col0}{ini}:{col1}{ini + len(mat) - 1}").Value = mat

        libro.Save()
        print(f"  [OK] Excel actualizado: R1={len(r1)}, R2={len(rango2)}, "
              f"R3={len(r3)}, R4={len(r4)}")
    finally:
        if libro is not None:
            libro.Close(SaveChanges=True)
        excel.Quit()


def actualizar_openpyxl(r1, r3, r4, rango2, jira_filas):
    print("  [AVISO] Mac/openpyxl: puede tardar y perder gráficos. "
          "Se recomienda Windows para producción.")
    wb = openpyxl.load_workbook(ADICIONALES)
    wd = wb[HOJA_DATOS]
    # DATA IN
    if wd.max_row >= FILA_DATOS:
        wd.delete_rows(FILA_DATOS, wd.max_row - FILA_DATOS + 1)
    for i, fila in enumerate(jira_filas, start=FILA_DATOS):
        for j, v in enumerate(fila, start=1):
            wd.cell(row=i, column=j, value=_normalizar(v))

    maximo = max(len(r1), len(r3), len(r4), len(rango2), RANGO2_FILAS)
    extra = max(0, maximo - RANGO2_FILAS)
    wp = wb[HOJA_PLAN]
    if extra:
        wp.insert_rows(25, extra)
    fin = FILA_FIN + extra
    for cols in (("B", "G"), ("I", "N"), ("Q", "W"), ("Y", "AD")):
        for r in range(FILA_DATOS, fin + 1):
            for c in range(openpyxl.utils.column_index_from_string(cols[0]),
                           openpyxl.utils.column_index_from_string(cols[1]) + 1):
                wp.cell(row=r, column=c, value=None)

    def escribir(celda, filas, ncols):
        from openpyxl.utils import column_index_from_string, get_column_letter
        c0 = column_index_from_string(celda[0])
        r0 = int(celda[1:])
        for i, fila in enumerate(filas):
            for j in range(ncols):
                wp.cell(row=r0 + i, column=c0 + j,
                        value=_normalizar(fila[j]) if j < len(fila) else None)

    escribir("B3", r1, 6)
    escribir("I3", rango2, 6)
    escribir("Q3", r3, 7)
    escribir("Y3", r4, 6)
    wb.save(ADICIONALES)
    wb.close()
    print(f"  [OK] openpyxl actualizado: R1={len(r1)}, R2={len(rango2)}, "
          f"R3={len(r3)}, R4={len(r4)}")


def main():
    print("=" * 70)
    print("  FASE 8 · Adicionales")
    print("=" * 70)
    if not ADICIONALES.is_file():
        sys.exit(f"No existe {ADICIONALES}")
    if not JIRA.is_file():
        sys.exit(f"No existe {JIRA}")

    filas_jira, regs = leer_jira()
    print(f"  JIRA: {len(regs):,} tickets")
    fuente_r2 = BACKUP if BACKUP.is_file() else ADICIONALES
    rango2 = leer_rango2(fuente_r2)
    r1, r3, r4 = construir(regs, rango2)
    print(f"  Rangos calculados: R1={len(r1)}, R2={len(rango2)}, "
          f"R3={len(r3)}, R4={len(r4)}")
    if "--dry-run" in sys.argv:
        print("  [DRY-RUN] No se modifica ningún archivo.")
        return

    if not BACKUP.is_file():
        shutil.copy2(ADICIONALES, BACKUP)
        print(f"  Foto fija creada: {BACKUP.name}")
    else:
        print(f"  Foto fija existente: {BACKUP.name} (no se toca)")

    if os.name == "nt":
        actualizar_excel(r1, r3, r4, rango2, filas_jira)
    else:
        actualizar_openpyxl(r1, r3, r4, rango2, filas_jira)


if __name__ == "__main__":
    main()