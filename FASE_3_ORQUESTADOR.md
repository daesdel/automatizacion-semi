# Fase 3 — Orquestador `main.py`

> [← Volver al índice general](DOCUMENTACION_PROYECTO.md)

Ejecuta Fase 1, Fase 2 y Fase 5 en cadena:

1. **Preflight**: comprueba los archivos de entrada (base, legal, seguimiento,
   Revision y Hoja sin título) y los scripts; si falta algo, no ejecuta nada.
2. **Fase 1** con `--refresh` y rutas explícitas (genera `JIRA.xlsx`).
3. **Fase 2** con el CSV.
4. **Fase 5** con el archivo anterior (opcional), `JIRA.xlsx`, el legal fechado y
   `seguimiento.xlsx`.
5. **Resumen**: tiempos por fase, total y tres archivos generados.

`main.py` acepta `--datos=CARPETA`: fija la carpeta de datos para los valores
por defecto, el log `Carpeta de datos` y el `JIRA.xlsx` generado. La app le pasa
exactamente la ruta escrita por el usuario, de modo que Fase 1, Fase 6 y Fase 7
trabajan sobre esa misma carpeta.

Garantías:
- Si una fase falla, se detiene y **no lanza la siguiente**.
- Si la fase fallida dejó una salida a medias, la elimina.
- Guarda `stdout/stderr` de cada fase.

```bash
python automatizacion/main.py
python automatizacion/main.py --sin-refresh
python automatizacion/main.py --seguimiento=R --jira=R --hoja=R --csv=R --out=CARPETA
```

Tiempo de referencia: **≈60 s** (10 s + 8 s + 41 s en Mac).

---

