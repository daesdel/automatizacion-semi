# Documentación del proyecto — Automatización de Informes SEMI

> Índice general del proyecto. Cada fase tiene su documento específico.

---

## 1. Visión general

El proyecto automatiza la generación de los informes que antes se construían
ejecutando macros de Excel:

| Fase | Documento | Script | Qué produce |
|---|---|---|---|
| 1 | [Fase 1](FASE_1_SEGUIMIENTO.md) | `fase1_seguimiento.py` | `seguimiento.xlsx` + `JIRA.xlsx` |
| 2 | [Fase 2](FASE_2_LEGAL.md) | `fase2_legal.py` | `DDMMYYYY_legal ministerio.xlsx` |
| 3 | [Fase 3](FASE_3_ORQUESTADOR.md) | `main.py` | Ejecuta Fases 1 + 2 + 5 |
| 4 | [Fase 4](FASE_4_APP_WEB.md) | `app_streamlit.py` | Interfaz web |
| 5 | [Fase 5](FASE_5_PREVISIONES.md) | `fase5_previsiones.py` | `Seguimiento_Previsiones_Limpio.xlsx` |
| 6 | [Fase 6](FASE_6_AUDITORIA.md) | `fase6_auditoria.py` | `Auditoria_Tickets_Faltantes.xlsx` |
| 7 | [Fase 7](FASE_7_FICHERO.md) | `fase7_fichero.py` | `DDMMYYYY_FICHERO.xlsx` |
| 8 | [Fase 8](FASE_8_ADICIONALES.md) | pendiente | `ADICIONALES SEGUIMIENTO.xlsx` |

```
semi/
├── DOCUMENTACION_PROYECTO.md
├── .streamlit/
│   └── config.toml                 # Streamlit: developmentMode = false
├── automatizacion/
│   ├── fase1_seguimiento.py
│   ├── fase1_seguimiento_v1.py     # copia histórica congelada (referencia)
│   ├── fase2_legal.py
│   ├── main.py
│   ├── app_streamlit.py
│   ├── fase5_previsiones.py
│   ├── data_in_headers.py           # cabeceras e índice de DATA IN
│   ├── previsiones_headers.py       # 87 cabeceras de PREVISIONES JUAN
│   └── estilos_previsiones.py       # especificación de formatos
├── subidas/                        # solo mientras corre la app (se vacía)
└── Nueva carpeta (2)/
    ├── Seguimiento E2E VODAFONE.xlsx
    ├── JIRA con DR.xlsx
    ├── Hoja de cálculo sin título.xlsx
    ├── MinisterioPowerBi.csv
    ├── Seguimiento_Previsiones_Limpio.xlsx   # base (semana anterior) y salida
    └── resultados/
        ├── seguimiento.xlsx
        └── DDMMYYYY_legal ministerio.xlsx
```

### Dependencias

```bash
pip install pandas openpyxl xlsxwriter python-calamine [streamlit] [psutil]
```

- `python-calamine`: lectura rápida de Excel (6× más rápida que openpyxl).
- `psutil`: opcional, solo para mostrar el pico de memoria.
- `streamlit`: solo para la interfaz (Fase 4).

---



## 2. Problemas, validación y Git

### 2.1 Problemas conocidos y soluciones

| Síntoma | Causa / solución |
|---|---|
| `ModuleNotFoundError: resource` | Resuelto: `rss_mb()` usa psutil → resource (Unix) → 0.0 |
| `UnicodeEncodeError` en Windows | `reconfigure(encoding="utf-8", errors="replace")` + mensajes ASCII |
| Fase 1 tarda 200 s en Windows | Falta `python-calamine`: `py -m pip install python-calamine` |
| Lee la hoja equivocada | Resuelto: selección dinámica por cobertura de columnas |
| Streamlit se cuelga con archivos en `%TEMP%` | Resuelto: carpeta local `subidas/` |
| `server.port does not work...` | `.streamlit/config.toml` con `developmentMode = false` |
| La app se colapsa en Windows | Modo carpeta local + `Popen` en streaming (sin buffer) y timeout 60 min |
| Fechas corruptas año 2525 | `agregar_fechas()` trabaja en microsegundos |

---

### 2.2 Historial de validación

- Fase 1: 100 % de las 121 columnas de datos idénticas; 99,8 %+ en fechas;
  VLOOKUP 100 % consistente en 300 filas aleatorias.
- Fase 2: 28.395/15.803 claves (coinciden con pivote_vdf/osp); 330/330
  celdas iguales.
- Fase 3: ejecución completa verificada y pruebas de fallo (preflight y
  limpieza de salidas parciales).
- Fase 4: prueba end-to-end con los 4 archivos reales (45,9 s), descargas
  disponibles y `subidas/` vacía al terminar.
- Fase 5: generación desde el archivo anterior probada en Mac; A/H/AB/AJ
  calculadas, fórmulas heredadas, DETALLE_TICKETS visible y 21 reglas de
  formato condicional.
- Fase 1: 8.501 filas originales vs 8.195 generadas; 306 finales vacías y
  7.649/7.649 IR conservados (`validar_irs`).
- Pipeline completo Fases 1+2+5: 67,89 s (9,99 + 8,11 + 49,80) y tres
  archivos generados, incluido el legal fechado `26102026_legal ministerio.xlsx`.
- Streamlit en modo carpeta local: 136,8 s el pipeline + 32,7 s la Fase 7,
  4 descargas (incluida `29102026_FICHERO.xlsx`), sin errores.
- `Carpeta de datos` del log = ruta exacta de la app (`/tmp/semi_test`).
- Auditoría: 21,1 s con descarga independiente.
- Fase 7 con RFOA corrupto: error amigable, sin Traceback.
- Fase 7: 27,9 s por CLI; 3 hojas, 41 filas RFOA visibles, 0 duplicados.
- Fase 6 con `JIRA.xlsx` correcto: 7.162 huérfanos; con archivo equivocado
  muestra error amigable, sin Traceback.
- Fase 5: `Informe_Cambios.xlsx` con 3 hojas; ZOOM y DETALLE excluyen AU
  informada y RIP; H/AB/AJ recalculadas en todas las filas.
- Fase 6: auditoría afinada a IR de radio con actividad; 10.844 elegibles,
  **7.162 huérfanos** la primera vez y **0 nuevos** al repetir (incremental).
- Fase 5: PREVISIONES y ZOOM conservan sus fórmulas (verificado C2/O2/B8).
- Fase 5 rellena con Translator las fórmulas que falten en filas antiguas,
  no solo en las nuevas.
- Sanity con calamine: 8,2 s para JIRA + Hoja + DATA IN.
- Fase 7: 3 hojas con valores, sin fórmulas ni hojas ocultas (1,55 MB, 31 s
  usando cachés de Excel). RFOA A1 con color negro `FF000000` y fondo blanco.
- Fase 1 usa xlsxwriter con formatos por columna (cabecera/dato) del E2E
  original: archivo válido, 8.194 filas, 5,29 MB, 23 s. Se evita el guardado
  openpyxl que corrompía el libro.
- pycel: bloque de 50 filas calculado en 8,7 s; H2 coincide con el valor de
  Excel (`Capacidad (Interna) 5G 700;NR700 EA`).
- RFOA SIN PREVISION: 41 filas visibles y estilos clonados de `Nodos`.
- ZOOM: 7 cuadros, cabeceras `WXX_AAAA`, columna `> 5 sem`, exclusión de
  ejecutadas antes y cuadro final SIN ESCENARIO.
- DETALLE_TICKETS visible, una fila por IR (9.721) y escenarios concatenados.
- Sin plantilla: PREVISIONES se hereda del archivo de la semana anterior.
- Git: repositorio actualizado; ver el último commit en `git log`.

---

### 2.3 Publicar el repositorio en GitHub

El repositorio local ya está inicializado y con el primer commit:

```bash
git log --oneline -1
# (hash del último commit)
```

`.gitignore` usa **lista blanca**: solo se versionan los `.py`, este
`DOCUMENTACION_PROYECTO.md` y el propio `.gitignore`. Quedan fuera
`__pycache__/`, `subidas/`, `.venv/`, los Excel/CSV, PDF, PPTX, HTML, etc.

### Conectar con GitHub

1. Crea un repositorio vacío en GitHub (sin README ni .gitignore).
2. Desde la carpeta del proyecto:

```bash
git remote add origin https://github.com/daesdel/automatizacion-semi.git
git branch -M main
git push -u origin main
```

Si usas SSH:

```bash
git remote add origin git@github.com:daesdel/automatizacion-semi.git
git branch -M main
git push -u origin main
```

Notas:

- Si GitHub pide contraseña, usa un **Personal Access Token** como contraseña
  (o `gh auth login` si tienes GitHub CLI).
- Los Excel/CSV (incluido el archivo anterior y la salida) **no** se suben
  (excluidos por `.gitignore`); cópialos aparte o guárdalos en OneDrive.
- Para actualizar el repo tras cambios:

```bash
git add .
git commit -m "Descripcion del cambio"
git push
```
