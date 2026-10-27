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

### 2.7 Control de filas perdidas

Investigación sobre las últimas filas de `NUEVO SEGUIMIENTO`:

- Original: **8.501 filas**; generado: **8.195 filas**. Las **306 filas
  restantes están completamente vacías** (ni IR ni datos en A:EI).
- Los **7.649 IR son idénticos** en orden, conjunto y último valor
  (`IR-120228`); no falta ningún ticket.
- `dropna(how="all")` se aplica solo a `VOLCADO JIRA` y `Encendidos ZR`,
  **nunca** a `NUEVO SEGUIMIENTO`.
- Nueva validación `validar_irs(df_ns, df_out)`: compara los IR de origen y
  salida y **aborta** si falta alguno, mostrando cuántos hay y el último.

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

1. **Preflight**: comprueba los archivos de entrada (base, legal, seguimiento,
   Revision y Hoja sin título) y los scripts; si falta algo, no ejecuta nada.
2. **Fase 1** con `--refresh` y rutas explícitas (genera `JIRA.xlsx`).
3. **Fase 2** con el CSV.
4. **Fase 5** con el archivo anterior (opcional), `JIRA.xlsx`, el legal fechado y
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
   4 archivos base de ahí y escribe `resultados_app/` en la misma carpeta; no
   sube nada ni usa `subidas/`.
2. **Subida**: cuatro `st.file_uploader` obligatorios (Seguimiento, JIRA con
   DR, Hoja sin título y CSV del Ministerio).

En modo carpeta se autodetecta `Seguimiento_Previsiones_Limpio.xlsx` en esa
ruta, y hay un campo extra para indicar otra ubicación. Se puede aportar
opcionalmente el
`Seguimiento_Previsiones_Limpio.xlsx` de la **semana anterior** para heredar
las fórmulas de PREVISIONES JUAN. Si no se aporta, se genera desde cero
**sin fórmulas** y se avisa en la interfaz.

En ambos modos:

- Botón **🚀 Ejecutar Proceso**.
- **Cronómetro** con fecha/hora de inicio y tiempo transcurrido.
- **Barra de progreso real** por fases (1, 2 y 5).
- **Consola en vivo** con las últimas líneas de `print`.
- Botón **🔎 Ejecutar Auditoría de Tickets** (Fase 6).
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

### 6.1 Arquitectura (sin plantilla)

La Fase 5 ya **no usa plantilla ni el libro original**. Usa como base el
`Seguimiento_Previsiones_Limpio.xlsx` de la **semana anterior**:

| Pieza | Papel |
|---|---|
| Archivo anterior | `PREVISIONES JUAN` con sus fórmulas y filas; se heredan y se añaden IRs nuevos |
| `data_in_headers.py` | Cabeceras (205 con nombre + 56 vacías) e índice 1…261 de `DATA IN` |
| `previsiones_headers.py` | 87 cabeceras de PREVISIONES JUAN (para arranque desde cero) |
| `estilos_previsiones.py` | Especificación estática de fuentes, rellenos, bordes, anchos, alto y freeze |
| `fase5_previsiones.py` | Calcula A/H/AB/AJ, hereda fórmulas y genera el libro final |

Origen de cada hoja del libro final:

| Hoja | Fuente |
|---|---|
| `DATA IN` | `data_in_headers.py` + datos de `JIRA.xlsx` (`IR_VDF`, desde fila 5) |
| `LEGAL VDF` / `LEGAL OSP` | `resultados/DDMMYYYY_legal ministerio.xlsx` (hojas `legal_vdf` / `legal_osp`) |
| `PREVISIONES JUAN` | Archivo anterior (filas + fórmulas) + IRs nuevos de `drive` + A/H/AB/AJ calculadas |
| `drive` | `resultados/seguimiento.xlsx` columnas B, AC, BV, DY |
| `Leg_Wn` | `Revision Legalizacion SEMI SemN.xlsx` → `Detalle` (A:I, M, V:Z) |
| `ZOOM` | Se calcula con `COUNTIFS` |
| `DETALLE_TICKETS` | Drill-down **visible**, una fila por IR, generado desde `JIRA.xlsx` |

### 6.2 Fórmulas heredadas y ArrayFormula

Las fórmulas se copian del archivo anterior. Para las **filas nuevas** se toma
la fórmula de la última fila existente y se traslada con
`openpyxl.formula.translate.Translator`, ajustando referencias relativas.

En la columna **AL** (“Datos Radio Enviados”) hay **3.775 fórmulas
matriciales** (`<f t="array" ref="AL2">`). openpyxl las devuelve como objetos
`ArrayFormula`, no como texto; se escribe su `.text` con
`write_array_formula`.

Solución aplicada:

```python
if isinstance(valor, ArrayFormula):
    ws.write_array_formula(fila, col, fila, col, valor.text, fmt)
```

Además, xlsxwriter convertía automáticamente `XLOOKUP`/`SWITCH` en matriciales
(45.000 celdas). Se neutraliza con `FormulasNormales`, de modo que **solo las
3.775 celdas de AL** quedan como matriciales y el resto son fórmulas normales.

### 6.3 Formato

El formato se aplica desde `estilos_previsiones.py` (extraído una vez del
libro original):

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
- Base común: patrón + `Fecha PO` (T) > 0 + Estado distinto de `RIP`.
- **Exclusión**: no se cuentan tickets con Fecha Real RFOA (AU) informada.
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

### 6.5 Columnas manuales y DETALLE_TICKETS

- **A (IR)**: IRs del archivo anterior + los que aparecen en `drive` después
  del último IR anterior localizado (si el último ya no está, se busca hacia
  atrás el último presente).
- **H (Clasificación)**: si el ticket encaja en uno o varios de los 6
  escenarios, se concatenan sus nombres con ` + `; si no encaja en ninguno,
  se escribe `Proyecto & Subproyecto`.
- **AB (IR-Tx)**: `Subtareas` de `Hoja sin título`: filas con
  `Subtareas.issueType = "Tx Single RAN"` y misma clave; se toma
  `Subtareas.issueKey`. Si no hay, `Sin IR`.
- **AJ (IR-Activación)**: si `OSP Tecnología IR` está vacío/`-`/`N/A`, `N/A`.
  Si hay tecnologías, se busca `ACTIVACION OR` en `Subtareas` y se toma
  `Subtareas.issueKey`; si no hay, `Sin IR`.

`DETALLE_TICKETS` es **visible**, con **una fila por IR** y los escenarios
concatenados con ` + `. Excluye los tickets con AU informada y los que están
en estado `RIP`. La semana de previsión muestra `Sin previsión` cuando AV está
vacía, o cuando AV es pasada y **no** hay Fecha RFOA real (AU).

### 6.6 Informe de cambios

Cada ejecución genera `Informe_Cambios.xlsx` junto al libro final, con:

- `ESCENARIO_CAMBIADO`: IR, H anterior y H actual.
- `IR_TX_AJ_ACTUALIZADO`: IR, AB/AJ anterior y actual (cuando antes no había
  IR y ahora sí).
- `NUEVOS`: IRs añadidos desde `drive`.

### 6.7 Fase 6 — Auditoría de tickets

`fase6_auditoria.py` compara con la columna A de `PREVISIONES JUAN` solo los
**IR de radio** (`IR-...`) que estén en `drive` **o** tengan `Fecha PO > 0` en
`DATA IN`. Excluye DR, RE y tickets sin actividad. Genera
`Auditoria_Tickets_Faltantes.xlsx` con IR, Origen, H, AB y AJ.

La app incorpora el botón **🔎 Ejecutar Auditoría de Tickets** y su descarga.

### 6.8 Auditoría de las macros 7, 8 y 9

- **copiar_juan (Módulo8)**: crea `DD-MM-YYYY_FICHERO.xlsx`; copia
  `SEGUIMIENTO PREVISIONES.xlsx` → `PREVISIONES JUAN` (A:CA, valores y
  formatos) y la tabla `Tabla1` de `Previsiones de RFOA_central.xlsx` (hoja
  `Nodos`, solo celdas visibles) → `RFOA SIN PREVISION`; añade `DUPLICADOS`.
- **copiar_ADICIONALES (Módulo3)**: `JIRA.XLSX` → hoja `IR_VDF`, rango
  `A5:GX60000`, pegado como valores en `ADICIONALES SEGUIMIENTO.xlsx` →
  `DATA IN!A3`.
- **copiar_GAPSEMI (Módulo4)**: `JIRA.XLSX` → `IR_VDF`, `A5:GX60000` → valores
  en `GAP semi.xlsx` → `DATA IN!A3` (limpia filtros antes).
- **copiar_DRIVE_SEMI (Módulo6)**: `Seguimiento.XLSX` → `Nuevo seguimiento`,
  `A5:DY10000` → valores en `GAP semi.xlsx` → `SEGUIMIENTO!A4`.

**Riesgos detectados para las Fases 7–9** (no repetir en Python):

- Rutas absolutas hardcodeadas (`C:\Users\describa\Downloads\Nueva carpeta (2)\`).
- Rangos fijos (`A:CA`, `GX60000`, `DY10000`) en vez de última fila/columna real.
- `PasteSpecial` sin comprobar filas vacías, filtros ni duplicados.
- Nombres de archivo/pestaña sensibles a mayúsculas y a la fecha del sistema.
- En Python: usar `pathlib`, rangos dinámicos (`max_row`/`max_column`) y
  validaciones de existencia antes de abrir.

### 6.9 Generación

```bash
# cada actualización
python automatizacion/fase5_previsiones.py [--anterior=RUTA]
# sin --anterior (o si no existe) se genera desde cero, sin fórmulas
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
- Fase 5: generación desde el archivo anterior probada en Mac; A/H/AB/AJ
  calculadas, fórmulas heredadas, DETALLE_TICKETS visible y 21 reglas de
  formato condicional.
- Fase 1: 8.501 filas originales vs 8.195 generadas; 306 finales vacías y
  7.649/7.649 IR conservados (`validar_irs`).
- Pipeline completo Fases 1+2+5: 67,89 s (9,99 + 8,11 + 49,80) y tres
  archivos generados, incluido el legal fechado `26102026_legal ministerio.xlsx`.
- Streamlit en modo carpeta local con archivo anterior: 137,3 s, 3 descargas,
  cronómetro/progreso/consola y botón de auditoría, sin errores.
- Fase 5: `Informe_Cambios.xlsx` con 3 hojas; ZOOM y DETALLE excluyen AU
  informada y RIP; H/AB/AJ recalculadas en todas las filas.
- Fase 6: auditoría afinada a IR de radio con actividad; 10.844 elegibles y
  **7.162 huérfanos** en la prueba, con H/AB/AJ calculadas.
- ZOOM: cabeceras `WXX_AAAA`, columna `> 5 sem`, exclusión de ejecutadas antes.
- DETALLE_TICKETS visible, una fila por IR (9.721) y escenarios concatenados.
- Sin plantilla: PREVISIONES se hereda del archivo de la semana anterior.
- Git: repositorio actualizado; ver el último commit en `git log`.

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
- Los Excel/CSV (incluido el archivo anterior y la salida) **no** se suben
  (excluidos por `.gitignore`); cópialos aparte o guárdalos en OneDrive.
- Para actualizar el repo tras cambios:

```bash
git add .
git commit -m "Descripcion del cambio"
git push
```
