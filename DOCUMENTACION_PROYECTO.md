# Documentación del proyecto — Automatización de Informes SEMI

> Manual técnico de las fases 1 a 4 (implementadas y validadas) y anexo de
> análisis de la fase 5 (pendiente de implementar).

---

## 1. Visión general

El proyecto automatiza la generación de los informes que antes se construían
ejecutando macros de Excel:

| Fase | Script | Qué produce |
|---|---|---|
| 1 | `automatizacion/fase1_seguimiento.py` | `resultados/seguimiento.xlsx` |
| 2 | `automatizacion/fase2_legal.py` | `resultados/DDMMYYYY_legal ministerio.xlsx` |
| 3 | `automatizacion/main.py` | Ejecuta Fases 1 + 2 + 5 en cadena |
| 4 | `automatizacion/app_streamlit.py` | Interfaz web para no usar la terminal |
| 5 | `automatizacion/fase5_previsiones.py` | `Seguimiento_Previsiones_Limpio.xlsx` |

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
│   ├── crear_plantilla.py           # extrae la plantilla (una vez)
│   └── estilos_previsiones.py       # especificación de formatos
├── subidas/                        # solo mientras corre la app (se vacía)
└── Nueva carpeta (2)/
    ├── Seguimiento E2E VODAFONE.xlsx
    ├── JIRA con DR.xlsx
    ├── Hoja de cálculo sin título.xlsx
    ├── MinisterioPowerBi.csv
    ├── plantilla_previsiones.xlsx            # plantilla Fase 5
    ├── Seguimiento_Previsiones_Limpio.xlsx   # salida Fase 5
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

## 2. Fase 1 — Seguimiento E2E

### 2.1 Qué hace

Actualiza el seguimiento semanal:

1. Lee `Seguimiento E2E VODAFONE.xlsx` (hojas `NUEVO SEGUIMIENTO`,
   `VOLCADO JIRA`, `Encendidos ZR`).
2. Con `--refresh`, reconstruye `VOLCADO JIRA` desde:
   - `JIRA con DR.xlsx` → se copia por encabezados a `Hoja sin título` (hoja
     con más columnas válidas).
   - `Hoja sin título` (Hoja 2 + Hoja 1 + ing pdte) → consolidado `IR_VDF` en
     memoria.
   - `IR_VDF` → `VOLCADO JIRA` por clave.
3. Recalcula las fórmulas de `NUEVO SEGUIMIENTO` de forma vectorizada.
4. Escribe una única hoja con **valores** (sin fórmulas):
   `resultados/seguimiento.xlsx`.

### 2.2 Entradas / salidas

| Entrada | Uso |
|---|---|
| `Seguimiento E2E VODAFONE.xlsx` | Libro base (8.191 × 139) |
| `JIRA con DR.xlsx` | Export JIRA (≈37.000 × 214) |
| `Hoja de cálculo sin título.xlsx` | Consolidado JIRA (Hoja 2, Hoja 1, ing pdte) |

| Salida | Contenido |
|---|---|
| `resultados/seguimiento.xlsx` | Hoja `NUEVO SEGUIMIENTO`, 8.191 × 139, solo valores |
| `JIRA.xlsx` (carpeta de datos) | Hoja `IR_VDF`, 261 columnas, datos desde fila 5; se regenera con `--refresh` |

### 2.3 Fórmulas replicadas

| Columnas | Lógica |
|---|---|
| M, N, S, EA, EG, EI | `VLOOKUP(clave, VOLCADO JIRA)` → WO NCR1, Estado, Hot Swap, Vendor, PIR, TX Ready |
| CT, CU | `VLOOKUP(clave, Encendidos ZR)` → ENVIADO, PLANTILLAS RECIBIDAS |
| CZ, DD, DK | `WEEKNUM(fecha)` sistema 1 (semana empieza en domingo) |
| DM | `IF(NCR-2<>"", IF(NCR-3<>"", "LCC DISEÑO ACABADO", "PENDIENTE DISEÑO RADIO"), "")` |
| DP–DU | `IFERROR(TEXT(LARGE(rango,1),"DD/MM/YYYY"), JOIN(" - ", rango))` |

### 2.4 Decisiones técnicas clave

- **calamine** como motor de lectura (`engine="calamine"`); si no está
  instalado, avisa y cae a openpyxl (mucho más lento).
- **`keep_default_na=False`**: el libro usa `"NA"`/`"N/A"` como valores de
  negocio (54.392 celdas). Tratarlos como nulos rompía los JOIN/VLOOKUP.
- **Encabezados dinámicos**: se busca la fila de encabezados (primeras 30
  filas, ≥5 celdas y que contenga «clave»), no se asume la fila 1.
- **Hojas dinámicas**: se elige la hoja con más columnas requeridas
  (preferencia `Hoja 2`), sin depender del nombre fijo.
- **Máximo de fechas en microsegundos**: `DataFrame.max(axis=1)` provoca
  segfault en pandas 3.0.6 con fechas de unidades mixtas (hay años 2525 fuera
  del rango de nanosegundos).
- **Escritura `openpyxl` write_only**: solo la hoja necesaria, memoria
  constante. (El `constant_memory` de xlsxwriter está roto en esta versión.)
- **Refresco en memoria**: no se escribe un `JIRA.xlsx` intermedio en disco.

### 2.5 Export `JIRA.xlsx` (para la Fase 5)

Con `--refresh` (y salvo `--no-export-jira`) se regenera
`Nueva carpeta (2)/JIRA.xlsx`:

- **Esquema**: el del `JIRA.xlsx` anterior (261 columnas), o el de la hoja
  con más columnas si el archivo no existe.
- **Base**: `Hoja sin título` (todas sus filas, ~32.400).
- **Actualización**: se sobrescriben las columnas coincidentes con
  `JIRA con DR` (que puede ser un subconjunto, ~29.800 filas).
- **Estructura**: fila 1 `VODAFONE`, fila 2 `IR_COMPLETO (Fiber2U)`,
  fila 3 vacía, fila 4 cabeceras y datos desde la fila 5 (lo que espera la
  macro `copiar_jira`).
- Escritura en streaming: solo celdas con valor (~39 s adicionales).

### 2.6 Ejecución

```bash
# Flujo completo (refresca VOLCADO JIRA)
python automatizacion/fase1_seguimiento.py --refresh

# Solo recálculo
python automatizacion/fase1_seguimiento.py

# Opciones
--seguimiento=RUTA --jira=RUTA --hoja=RUTA
--out=CARPETA_SALIDA
--filas=N          # prueba con las primeras N filas
```

Rendimiento de referencia (Mac): **≈35 s** con refresh, ≈9 s sin refresh;
pico ≈0,9–1,3 GB.

---

## 3. Fase 2 — Legal (Ministerio VDF/OSP)

### 3.1 Qué hace

1. Lee `MinisterioPowerBi.csv` (93.836 × 30, separador coma, cabeceras entre
   corchetes).
2. Limpia fechas: elimina `T00:00:00` y convierte seriales Excel (época
   1899-12-30); conserva textos como `SI`/`NO`.
3. Reordena columnas `Y:AA → Q` (igual que la macro `ProcesarCSV`).
4. Separa por `Operador Responsable`: **VDF = Vodafone**, **OSP = Orange**.
5. Construye el equivalente a `Procesado_vdf`/`Procesado_osp`:
   `clave + tecnologia + EMF + AMT + PSC` y la clave compuesta
   `clave_tecnologia` (primera coincidencia, como VLOOKUP).
6. Construye la matriz final: 1 fila por clave única y **33 celdas**
   (11 tecnologías × 3 bloques EMF/AMT/PSC).
7. Escribe `resultados/DDMMYYYY_legal ministerio.xlsx` con las hojas
   `legal_vdf` y `legal_osp`, solo valores y fechas reales (`dd/mm/yyyy`).

### 3.2 Entradas / salidas

| Entrada | Salida |
|---|---|
| `MinisterioPowerBi.csv` | `resultados/DDMMYYYY_legal ministerio.xlsx` (hojas `legal_vdf` / `legal_osp`) |

### 3.3 Estructura de cada hoja

| Fila | Contenido |
|---|---|
| 1 | Números 1…37 |
| 2 | `EMF VDF` (B:L), `AMT VDF` (M:W), `PSC VDF` (X:AH) |
| 3 | `[Clave de incidencia]` + 11 tecnologías ×3 + `EMF`, `AMT`, `PSC` (AI:AK) |
| 4+ | Clave + 33 celdas + 3 columnas consolidadas |

`AI:AK` replican las fórmulas del export legal:
`=MAX(B:L)` (EMF), `=MAX(M:W)` (AMT) y `=MAX(X:AH)` (PSC). Si un bloque no
tiene fechas, se escribe `0` (como Excel). Validado 100 % contra
`13102026_legal ministerio.xlsx` (300 claves × 37 columnas por hoja).

Tecnologías: G900, U900, U2100, LTE800, LTE900, LTE1800, LTE2100, LTE2600,
5G 700, 5G 3700, 5G 2100.

### 3.4 Validación de referencia

- VDF: **28.395** claves (idéntico a `pivote_vdf`).
- OSP: **15.803** claves (idéntico a `pivote_osp`).
- 330/330 celdas comparadas contra el libro original.
- Tiempo ≈6 s.

```bash
python automatizacion/fase2_legal.py [--csv=RUTA] [--out=CARPETA]
```

---

## 4. Fase 3 — Orquestador `main.py`

Ejecuta Fase 1, Fase 2 y Fase 5 en cadena:

1. **Preflight**: comprueba los archivos de entrada (incluida la plantilla y
   la Revision) y los scripts; si falta algo, no ejecuta nada.
2. **Fase 1** con `--refresh` y rutas explícitas (genera `JIRA.xlsx`).
3. **Fase 2** con el CSV.
4. **Fase 5** con la plantilla, `JIRA.xlsx`, el legal fechado y
   `seguimiento.xlsx`.
5. **Resumen**: tiempos por fase, total y tres archivos generados.

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

## 5. Fase 4 — Interfaz Streamlit

`app_streamlit.py` ofrece **dos modos**:

1. **Carpeta local (recomendado en Windows)**: campo de texto con la ruta
   (p. ej. `C:\Users\describa\Downloads\Nueva carpeta (2)`). La app lee los
   5 archivos de ahí y escribe `resultados_app/` en la misma carpeta; no sube
   nada ni usa `subidas/`.
2. **Subida**: cinco `st.file_uploader` (Seguimiento, JIRA con DR, Hoja sin
   título, CSV del Ministerio y `plantilla_previsiones.xlsx`). Se guardan en
   `subidas/` con nombres limpios y se borran al terminar.

En ambos modos:

- Botón **🚀 Ejecutar Proceso**.
- Ejecuta **`main.py`** (Fases 1 + 2 + 5).
- Salida leída **en tiempo real** con `Popen` + cola y `PYTHONUNBUFFERED=1`
  (evita el bloqueo del buffer en Windows), con `st.status`, `st.progress` y
  un **timeout de 60 minutos**.
- `--jira-out` deja el `JIRA.xlsx` intermedio en la carpeta de salida (local),
  no en OneDrive.
- Al terminar: tiempo y **tres descargas**: `seguimiento.xlsx`,
  `DDMMYYYY_legal ministerio.xlsx` y `Seguimiento_Previsiones_Limpio.xlsx`.
- El pie de página menciona las Fases 1, 2 y 5.

### Arranque

```bash
# macOS (entorno de desarrollo con librerías en /tmp/my_py)
PYTHONPATH=/tmp/my_py python3 -m streamlit run automatizacion/app_streamlit.py

# Windows
py -m streamlit run automatizacion\app_streamlit.py
```

Navegador: <http://localhost:8501>

### Notas Windows/Mac

- `.streamlit/config.toml` fija `developmentMode = false` (necesario al
  instalar Streamlit fuera de `site-packages`).
- `subidas/` evita los bloqueos de OneDrive/Defender sobre `%TEMP%`.
- Si el puerto está ocupado: `--server.port 8502`.
- Sin `python-calamine`, Fase 1 avisa y tarda mucho más.

---

## 6. Fase 5 — Archivo limpio (`Seguimiento_Previsiones_Limpio.xlsx`)

### 6.1 Arquitectura de plantilla (sin dependencia del original)

La Fase 5 ya **no lee** `SEGUIMIENTO PREVISIONES.xlsx`. Se apoya en:

| Pieza | Papel |
|---|---|
| `plantilla_previsiones.xlsx` | **Solo** `PREVISIONES JUAN`: cabeceras, fórmulas, columnas manuales (A, H, AB, AJ) y celdas `AUTO` donde no hay fórmula |
| `crear_plantilla.py` | Genera la plantilla **una sola vez** desde el libro original |
| `data_in_headers.py` | Cabeceras (205 con nombre + 56 vacías) e índice 1…261 de `DATA IN` |
| `estilos_previsiones.py` | Especificación estática de fuentes, rellenos, bordes, anchos, alto y freeze |
| `fase5_previsiones.py` | Inyecta los datos de Fase 1/2 y genera el libro final |

Origen de cada hoja del libro final:

| Hoja | Fuente |
|---|---|
| `DATA IN` | `data_in_headers.py` + datos de `JIRA.xlsx` (`IR_VDF`, desde fila 5) |
| `LEGAL VDF` / `LEGAL OSP` | `resultados/DDMMYYYY_legal ministerio.xlsx` (hojas `legal_vdf` / `legal_osp`) |
| `PREVISIONES JUAN` | Plantilla (fórmulas + manuales; las celdas `AUTO` se dejan vacías) |
| `drive` | `resultados/seguimiento.xlsx` columnas B, AC, BV, DY |
| `Leg_Wn` | `Revision Legalizacion SEMI SemN.xlsx` → `Detalle` (A:I, M, V:Z) |
| `ZOOM` | Se calcula con `COUNTIFS` |
| `DETALLE_TICKETS` | Hoja **oculta** de drill-down generada desde `JIRA.xlsx` |

### 6.2 Arreglo de ArrayFormula

En la columna **AL** (“Datos Radio Enviados”) hay **3.775 fórmulas
matriciales** (`<f t="array" ref="AL2">`). openpyxl las devuelve como objetos
`ArrayFormula`, no como texto; al escribirlas con `str()` aparecía
`<openpyxl...ArrayFormula object at 0x...>`.

Solución aplicada:

```python
if isinstance(valor, ArrayFormula):
    ws.write_array_formula(fila, col, fila, col, valor.text, fmt)
```

Además, xlsxwriter convertía automáticamente `XLOOKUP`/`SWITCH` en matriciales
(45.000 celdas). Se neutraliza con `FormulasNormales`, de modo que **solo las
3.775 celdas de AL** quedan como matriciales y el resto son fórmulas normales.

### 6.3 Formato clonado

La plantilla arrastra el formato exacto del original:

- Cabecera con sus 19 variantes de relleno/fuente (rojo, amarillo, naranja,
  azul, morado…), bordes finos/medios y alto de fila 45,75.
- Datos con borde fino izquierda/derecha/abajo, sin relleno, Candara/Calibri
  según columna y formatos de fecha.
- Anchos de columna, freeze en B2 y 7 reglas de formato condicional
  (verde `#92D050`, gris `#8D8B8B` para RIP, rojo `#FF0000`).

### 6.4 ZOOM: 6 escenarios

| # | Campo | Patrón |
|---|---|---|
| 1 | Proyectos Radio (G) | `*portadoras) 5G 7*` |
| 2 | Proyectos Radio (G) | `*portadoras) 5g 37*` |
| 3 | Proyectos Radio (G) | `*portadoras*18*` (los datos reales son “Portadoras 5G 1800”) |
| 4 | Proyectos Radio (G) | `*4g*` |
| 5 | Proyecto & Subproyecto (E) | `*upgrade*` |
| 6 | Proyecto & Subproyecto (E) | `*jumping*` |

- 5 semanas dinámicas (lunes actual + 4) + No planificada + **> 5 sem** + Total.
- Cabeceras **`WXX_AAAA`** con `WEEKNUM(fecha,21)` (semana ISO) y año ISO.
  Los `COUNTIFS` usan una **fila oculta** con las fechas reales.
- Base común: patrón + `Fecha PO` (T) > 0.
- **Ejecutadas antes**: si la Fecha RFOA real (AU) es anterior al lunes actual,
  la incidencia **no** cuenta como planificada en ninguna semana (aunque su
  Previsión RFOA sea de esta semana o futura). Si AU es de esta semana, sí
  cuenta.
- `No planificada`: sin Fecha RFOA (AU vacía o `"-"`) **y** alguna de:
  - Previsión RFOA (AV) vacía o `"-"`,
  - AV anterior al lunes de la semana actual.
- **> 5 sem**: AV posterior al domingo de la semana 5 (y no ejecutada antes).
- **Total = 5 semanas + No planificada + > 5 sem**.
- Títulos combinados A:I, centrados, Calibri 14:
  `PORTADORAS 5G700`, `PORTADORAS 5G3700`, `PORTADORAS 5G1800`,
  `PORTADORAS 4G`, `CAPACIDAD UPGRADE`, `JUMPING`.
- Fila de parámetros (patrón/campo) **oculta**.
- Texto de la fila de datos: **“Nº de Tickets JIRA”**.

> **Nota sobre fórmulas “en español”**: el XML de un `.xlsx` guarda siempre
> los nombres de función en inglés (`TEXT`, `YEAR`, `WEEKDAY`, `WEEKNUM`).
> Escribir `TEXTO`, `AÑO`, `DIASEM` o `ISO.NUM.DE.SEMANA` dentro del archivo lo
> corrompe. Excel en español los muestra traducidos automáticamente. El error
> anterior venía de `ISOWEEKNUM`; se sustituyó por `WEEKNUM(fecha,21)`, que es
> nativo y funciona en Excel ES (se ve como `NUM.DE.SEMANA`).

### 6.5 Generación

```bash
# una sola vez (ya generada)
python automatizacion/crear_plantilla.py

# cada actualización
python automatizacion/fase5_previsiones.py
```

Salida: `Nueva carpeta (2)/resultados/Seguimiento_Previsiones_Limpio.xlsx` (~15,5 MB).
La primera apertura recalcula las fórmulas de PREVISIONES JUAN.

## 7. Problemas conocidos y soluciones

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

## 8. Historial de validación

- Fase 1: 100 % de las 121 columnas de datos idénticas; 99,8 %+ en fechas;
  VLOOKUP 100 % consistente en 300 filas aleatorias.
- Fase 2: 28.395/15.803 claves (coinciden con pivote_vdf/osp); 330/330
  celdas iguales.
- Fase 3: ejecución completa verificada y pruebas de fallo (preflight y
  limpieza de salidas parciales).
- Fase 4: prueba end-to-end con los 4 archivos reales (45,9 s), descargas
  disponibles y `subidas/` vacía al terminar.
- Fase 5: generación probada en Mac (39,9 s, 15,1 MB), plantilla de 4,9 MB,
  3.775 únicas fórmulas matriciales en AL y 21 reglas de formato condicional.
- Pipeline completo Fases 1+2+5: 65,69 s (9,96 + 7,85 + 47,89) y tres
  archivos generados, incluido el legal fechado `26102026_legal ministerio.xlsx`.
- Streamlit en modo carpeta local: 134,0 s, 3 descargas (la del legal con su
  nombre fechado), sin errores.
- ZOOM: cabeceras `WXX_AAAA` con `WEEKNUM(...,21)`, columna `> 5 sem`,
  exclusión de ejecutadas antes y hoja oculta `DETALLE_TICKETS` (9.388 filas).
- Plantilla: solo `PREVISIONES JUAN` (sin `DATA IN`).
- Git: repositorio inicializado, commit `0d2b14c`, solo 9 archivos versionados.

---

## 9. Publicar el repositorio en GitHub

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
- La plantilla `plantilla_previsiones.xlsx` y los datos **no** se suben
  (excluidos por `.gitignore`); cópialos aparte o guárdalos en OneDrive.
- Para actualizar el repo tras cambios:

```bash
git add .
git commit -m "Descripcion del cambio"
git push
```
