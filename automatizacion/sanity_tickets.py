#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sanity check de tickets en segundo plano.

Cuenta los tickets de JIRA con DR (Hoja 1) y Hoja sin título (Hoja 2), los
compara con DATA IN del Seguimiento_Previsiones_Limpio anterior y escribe un
JSON con el resultado.
"""

import json
import sys
from pathlib import Path

from python_calamine import CalamineWorkbook


def _arg(nombre, defecto=None):
    pref = f"--{nombre}="
    for a in sys.argv[1:]:
        if a.startswith(pref):
            return Path(a.split("=", 1)[1]).expanduser()
    return defecto


JIRA = _arg("jira")
HOJA = _arg("hoja")
PREV = _arg("prev")
OUT = _arg("out")


def contar(path, hoja, col, cabecera=1):
    """Cuenta filas con clave usando calamine (rápido, sin abrir Excel)."""
    if not path or not path.is_file():
        return None
    try:
        wb = CalamineWorkbook.from_path(str(path))
        ws = wb.get_sheet_by_name(hoja)
        filas = ws.to_python(skip_empty_area=True)
        n = sum(1 for fila in filas[cabecera:]
                if len(fila) >= col and fila[col - 1] not in (None, ""))
        return n
    except Exception as e:  # noqa: BLE001
        print(f"  [AVISO] No se pudo contar {path.name}/{hoja}: {e}")
        return None


def main():
    print("Comprobando integridad de los archivos...")
    anterior = contar(PREV, "DATA IN", 1, cabecera=2) if PREV else None
    print(f"  DATA IN anterior: {anterior if anterior is not None else '—'}")
    j = contar(JIRA, "Hoja 1", 2)
    print(f"  JIRA con DR (Hoja 1): {j if j is not None else '—'}")
    h = contar(HOJA, "Hoja 2", 1)
    print(f"  Hoja sin título (Hoja 2): {h if h is not None else '—'}")

    actual = max([x for x in (j, h) if x is not None], default=None)
    bajada = bool(anterior and actual is not None and actual < anterior)
    datos = {"anterior": anterior, "actual": actual, "bajada": bajada,
             "jira": j, "hoja": h}
    if OUT:
        OUT.parent.mkdir(parents=True, exist_ok=True)
        OUT.write_text(json.dumps(datos), encoding="utf-8")
    print(f"Resultado: anterior={anterior} actual={actual} bajada={bajada}")
    print("[OK] Comprobación terminada")


if __name__ == "__main__":
    main()