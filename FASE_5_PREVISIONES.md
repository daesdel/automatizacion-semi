# Fase 5 — Archivo limpio (`Seguimiento_Previsiones_Limpio.xlsx`)

> [← Volver al índice general](DOCUMENTACION_PROYECTO.md)

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
| 7 | Inverso de 1–6 | **SIN ESCENARIO** |

- 5 semanas dinámicas (lunes actual + 4) + No planificada + **> 5 sem** + Total.
- Cabeceras **`WXX_AAAA`** con `WEEKNUM(fecha,21)` (semana ISO) y año ISO.
  Los `COUNTIFS` usan una **fila oculta** con las fechas reales.
- Base común: patrón + `Fecha PO` (T) > 0 + Estado distinto de `RIP`.
- **Exclusión**: no se cuentan tickets con Fecha Real RFOA (AU) informada.
- **JUMPING** excluye los tickets que pertenecen a PORTADORAS 5G700, 5G3700,
  5G1800 o 4G.
- **SIN ESCENARIO** (séptimo cuadro, al final): cuenta los tickets con Fecha PO
  > 0 que **no** cumplen ninguno de los 6 escenarios. Se implementa con
  `COUNTIFS` y exclusiones `<>` sobre G (cuatro patrones de portadoras) y E
  (`upgrade`, `jumping`), manteniendo semanas, No planificada, > 5 sem y Total.
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



### 6.10 Cálculo: Excel COM o pycel

> **Windows es el entorno oficial de producción.** El cálculo se hace con
> Excel COM (`pywin32`), que recalcula todo el libro en segundos y genera el
> FICHERO limpio con valores. Mac queda para desarrollo y pruebas de 50 filas;
> el fallback pycel es funcional pero puede tardar horas con el libro completo.

La Fase 5 intenta recalcular en **Windows con Excel COM**:

1. Abre el libro en una instancia invisible de Excel.
2. `CalculateFullRebuild()` + guardar.
3. **Conserva las fórmulas vivas** y solo deja los valores en caché. El libro
   intermedio (`Seguimiento_Previsiones_Limpio.xlsx`) mantiene fórmulas en
   PREVISIONES JUAN y ZOOM para poder heredarlas la semana siguiente.

Si no hay Excel (Mac, pywin32 ausente, Excel ocupado), **no se usa
LibreOffice**. Se usa **pycel 1.0b30** por bloques de 50 filas
(`calc_pycel.py`):

- Expande `LET` y reescribe `XLOOKUP` como `INDEX/MATCH`.
- Acota rangos completos (`A:FS`, `A:A`) a las dimensiones reales.
- Añade `WEEKNUM` y `SEARCH` a pycel y aplica parches de compatibilidad con
  Python 3.15/openpyxl.
- Calcula celda a celda y sustituye las fórmulas por valores.

La Fase 7 es la **única** que convierte fórmulas a valores: usa las cachés de
Excel si existen y, solo si faltan, recurre a pycel. Garantiza **siempre 3
hojas** (`PREVISIONES JUAN`, `RFOA SIN PREVISION`, `DUPLICADOS`), sin fórmulas
ni hojas ocultas.

Dependencias en Mac/Windows sin Excel:

```bash
py -m pip install pycel==1.0b30 "networkx==2.6.3" ruamel.yaml
```

El cálculo completo en Mac puede tardar horas (bloques de 50 filas); en
Windows con Excel es cuestión de segundos. La Fase 5 y la Fase 7 muestran el
aviso: “En Mac el cálculo de fórmulas es extremadamente lento. Se recomienda
usar Windows para generar el FICHERO final.”

`--sin-recalcular` omite Excel, pero Fase 7 seguirá calculando con pycel para
entregar valores.



### 6.11 Generación

```bash
# cada actualización
python automatizacion/fase5_previsiones.py [--anterior=RUTA]
# sin --anterior (o si no existe) se genera desde cero, sin fórmulas
```

Salida: `Nueva carpeta (2)/resultados/Seguimiento_Previsiones_Limpio.xlsx` (~15,5 MB).
La primera apertura recalcula las fórmulas de PREVISIONES JUAN.

