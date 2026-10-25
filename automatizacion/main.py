#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FASE 3 - Orquestador: actualización completa (Seguimiento + Legal)
==================================================================

Ejecuta en cadena:
  1) Fase 1  ->  resultados/seguimiento.xlsx   (fase1_seguimiento.py --refresh)
  2) Fase 2  ->  resultados/Legal_Final.xlsx   (fase2_legal.py)

Garantías:
  - Comprueba TODOS los archivos de entrada antes de empezar (fail-fast).
  - Si una fase falla, se detiene y no lanza la siguiente.
  - Si la fase fallida dejó un archivo a medias, lo elimina.
  - Al final muestra tiempos, archivos generados y un resumen único.

Uso:
    PYTHONPATH=/tmp/my_py python3 automatizacion/main.py
    PYTHONPATH=/tmp/my_py python3 automatizacion/main.py --sin-refresh
        (la Fase 1 recalcula sobre el VOLCADO JIRA actual, sin refrescarlo)
"""

import os
import subprocess
import sys
import time
from pathlib import Path

# Consola de Windows (cp1252): evita UnicodeEncodeError al imprimir símbolos.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = Path(__file__).resolve().parent
BASE = HERE.parent
DATA = BASE / "Nueva carpeta (2)"


def _arg_path(nombre: str, defecto: Path) -> Path:
    """Ruta opcional por argumento: --nombre=/ruta/al/archivo."""
    pref = f"--{nombre}="
    for a in sys.argv[1:]:
        if a.startswith(pref):
            return Path(a.split("=", 1)[1]).expanduser()
    return defecto


OUT_DIR = _arg_path("out", DATA / "resultados")

FASE1 = HERE / "fase1_seguimiento.py"
FASE2 = HERE / "fase2_legal.py"
FASE5 = HERE / "fase5_previsiones.py"

JIRA_EXPORT = _arg_path("jira-out", DATA / "JIRA.xlsx")
PLANTILLA = _arg_path("plantilla", DATA / "plantilla_previsiones.xlsx")
REVISION = _arg_path("revision", DATA / "Revision Legalizacion SEMI SemN.xlsx")

ENTRADAS = {
    "Seguimiento E2E": _arg_path("seguimiento", DATA / "Seguimiento E2E VODAFONE.xlsx"),
    "JIRA con DR": _arg_path("jira", DATA / "JIRA con DR.xlsx"),
    "Hoja sin título": _arg_path("hoja", DATA / "Hoja de cálculo sin título.xlsx"),
    "CSV Ministerio": _arg_path("csv", DATA / "MinisterioPowerBi.csv"),
    "Plantilla previsiones": PLANTILLA,
    "Revision legalizacion": REVISION,
}

SALIDAS = {
    "Fase 1 - Seguimiento": OUT_DIR / "seguimiento.xlsx",
    "Fase 2 - Legal": OUT_DIR / "Legal_Final.xlsx",
    "Fase 5 - Previsiones": OUT_DIR / "Seguimiento_Previsiones_Limpio.xlsx",
}

SIN_REFRESH = "--sin-refresh" in sys.argv


# ----------------------------------------------------------------------
# Entorno
# ----------------------------------------------------------------------
def preparar_entorno() -> dict:
    """Permite ejecutar sin PYTHONPATH si las dependencias están en /tmp/my_py."""
    env = os.environ.copy()
    vendor = Path("/tmp/my_py")
    if vendor.is_dir():
        actual = env.get("PYTHONPATH", "")
        if str(vendor) not in actual.split(os.pathsep):
            env["PYTHONPATH"] = str(vendor) + (os.pathsep + actual if actual else "")
        if str(vendor) not in sys.path:
            sys.path.insert(0, str(vendor))
    return env


ENV = preparar_entorno()


def tam(path: Path) -> str:
    return f"{path.stat().st_size / 1e6:.2f} MB"


def preflight() -> bool:
    """Comprueba scripts y entradas antes de ejecutar nada."""
    problemas = []
    for script in (FASE1, FASE2, FASE5):
        if not script.is_file():
            problemas.append(f"Falta el script: {script}")
    for nombre, ruta in ENTRADAS.items():
        if not ruta.is_file():
            problemas.append(f"Falta el archivo de entrada [{nombre}]: {ruta}")
        elif ruta.stat().st_size == 0:
            problemas.append(f"Archivo vacío [{nombre}]: {ruta}")
    if problemas:
        print("\n[ERROR] No se puede iniciar el proceso:")
        for p in problemas:
            print("   -", p)
        print("\nNo se ha generado ningún archivo.")
        return False
    return True


# ----------------------------------------------------------------------
# Ejecución de una fase
# ----------------------------------------------------------------------
def ejecutar_fase(nombre: str, script: Path, args: list, salida: Path,
                  timeout: int = 1800) -> float:
    """Ejecuta una fase como subproceso. Devuelve segundos o lanza RuntimeError."""
    print("\n" + "=" * 72)
    print(f"  >> {nombre}")
    print("=" * 72)

    existia = salida.is_file()
    mtime_previo = salida.stat().st_mtime if existia else 0.0

    t0 = time.time()
    try:
        proc = subprocess.run(
            [sys.executable, str(script), *args],
            cwd=str(BASE), env=ENV, capture_output=True, text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        _limpiar_salida(salida, existia, mtime_previo)
        raise RuntimeError(f"{nombre} excedió el tiempo máximo ({timeout}s)")
    dur = time.time() - t0

    # Salida de la fase, indentada para distinguirla
    for linea in (proc.stdout or "").rstrip().splitlines():
        print("   " + linea)
    if proc.returncode != 0:
        for linea in (proc.stderr or "").rstrip().splitlines():
            print("   " + linea)
        _limpiar_salida(salida, existia, mtime_previo)
        raise RuntimeError(f"{nombre} terminó con código {proc.returncode}")

    if not salida.is_file():
        raise RuntimeError(f"{nombre} terminó bien pero no generó {salida.name}")

    print(f"   - {nombre}: OK en {dur:.2f}s -> {salida.name} ({tam(salida)})")
    return dur


def _limpiar_salida(salida: Path, existia: bool, mtime_previo: float) -> None:
    """Elimina el archivo de salida si lo creó/modificó la fase fallida."""
    try:
        if salida.is_file():
            modificado = (not existia) or salida.stat().st_mtime > mtime_previo
            if modificado:
                salida.unlink()
                print(f"   [AVISO] Eliminado {salida.name} para no dejar un archivo a medias.")
    except OSError as e:
        print(f"   [AVISO] No se pudo eliminar {salida.name}: {e}")


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main() -> int:
    t_total = time.time()
    print("=" * 72)
    print("  ACTUALIZACIÓN COMPLETA  ·  Fase 1 (Seguimiento) + Fase 2 (Legal)")
    print("=" * 72)
    print(f"  Carpeta de datos:  {DATA}")
    print(f"  Carpeta de salida: {OUT_DIR}")
    print(f"  Modo Fase 1: {'recálculo' if SIN_REFRESH else 'refresco completo (--refresh)'}")

    if not preflight():
        return 1

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    args1 = [
        f"--seguimiento={ENTRADAS['Seguimiento E2E']}",
        f"--jira={ENTRADAS['JIRA con DR']}",
        f"--hoja={ENTRADAS['Hoja sin título']}",
        f"--jira-out={JIRA_EXPORT}",
        f"--out={OUT_DIR}",
    ]
    if not SIN_REFRESH:
        args1.append("--refresh")
    args2 = [
        f"--csv={ENTRADAS['CSV Ministerio']}",
        f"--out={OUT_DIR}",
    ]
    args5 = [
        f"--plantilla={PLANTILLA}",
        f"--revision={REVISION}",
        f"--jira={JIRA_EXPORT}",
        f"--legal={SALIDAS['Fase 2 - Legal']}",
        f"--seguimiento={SALIDAS['Fase 1 - Seguimiento']}",
        f"--out={SALIDAS['Fase 5 - Previsiones']}",
    ]

    tiempos = {}
    try:
        tiempos["Fase 1"] = ejecutar_fase(
            "FASE 1 · SEGUIMIENTO E2E", FASE1, args1,
            SALIDAS["Fase 1 - Seguimiento"],
        )
        tiempos["Fase 2"] = ejecutar_fase(
            "FASE 2 · LEGAL (Ministerio)", FASE2, args2,
            SALIDAS["Fase 2 - Legal"],
        )
        tiempos["Fase 5"] = ejecutar_fase(
            "FASE 5 · PREVISIONES (archivo limpio)", FASE5, args5,
            SALIDAS["Fase 5 - Previsiones"],
        )
    except RuntimeError as e:
        transcurrido = time.time() - t_total
        print("\n" + "=" * 72)
        print(f"  [ERROR] PROCESO DETENIDO tras {transcurrido:.2f}s")
        print(f"  Motivo: {e}")
        print("=" * 72)
        if "Fase 2" in tiempos:
            print("  Fases 1 y 2 completadas; la Fase 5 no se ejecutó.")
        elif "Fase 1" in tiempos:
            print(f"  La Fase 1 sí terminó: {SALIDAS['Fase 1 - Seguimiento'].name} está disponible.")
            print("  La Fase 2 no se ejecutó; Legal_Final.xlsx no se ha actualizado.")
        else:
            print("  No se ha generado ningún archivo nuevo.")
        return 1

    total = time.time() - t_total
    print("\n" + "=" * 72)
    print("  [OK] Proceso completado. Archivos listos en la carpeta resultados/")
    print("=" * 72)
    print(f"  Tiempo total: {total:.2f}s"
          f"   (Fase 1: {tiempos['Fase 1']:.2f}s | Fase 2: {tiempos['Fase 2']:.2f}s"
          f" | Fase 5: {tiempos['Fase 5']:.2f}s)")
    print("\n  Archivos generados:")
    for nombre, ruta in SALIDAS.items():
        if ruta.is_file():
            hora = time.strftime("%H:%M:%S", time.localtime(ruta.stat().st_mtime))
            print(f"    • {ruta.name:<22} {tam(ruta):>9}  ({hora})   [{nombre}]")
    print(f"\n  Ubicación: {OUT_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())