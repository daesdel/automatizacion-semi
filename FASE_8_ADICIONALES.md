# Fase 8 — Adicionales (`ADICIONALES SEGUIMIENTO.xlsx`)

> [← Volver al índice general](DOCUMENTACION_PROYECTO.md)

## 8.1 Reconocimiento del libro

Archivo: `Nueva carpeta (2)/ADICIONALES SEGUIMIENTO.xlsx` (17 MB).

| Pestaña | Filas × columnas | Estado | Contenido |
|---|---:|---|---|
| `DATA IN` | 37.030 × 261 | visible | Datos maestros de tickets, mismo esquema que JIRA |
| `DATA IN_ZN` | 8.672 × 198 | oculta | Subconjunto ZN de datos maestros |
| `PLANIFICACION` | 88 × 43 | visible | Los 4 rangos de tickets, resumen por proyecto/semana, KO y re-planificación |
| `LEGAL OSP` | 2.882 × 33 | visible | Datos legales OSP |
| `Estados JUMPING_ZN` | 3.875 × 45 | oculta | Estados y acumulados JUMPING ZN; 4 gráficos |
| `BQA` | 50 × 293 | visible | Auditorías/reparos BQA; 4 gráficos |

## 8.2 Los cuatro rangos de PLANIFICACION

Cabecera de bloque en la fila 1, subcabecera en la fila 2 y datos desde la fila 3 hasta la 24 (22 filas por bloque).

| Rango | Celdas | Cabecera | Contenido |
|---|---|---|---|
| 1 | `B1:G24` | `SEMANA N-1 (no plan)` | Tickets con RFOA de la semana anterior que ya no están en el rango 2 |
| 2 | `I1:O24` | `SEMANA N-1` | Tickets que la semana anterior tenían previsión de RFOA en esa semana |
| 3 | `Q1:W24` | `SEMANA N` | Tickets que la semana anterior tenían previsión de RFOA en la semana actual |
| 4 | `Y1:AD24` | `SEMANA N+1` | Tickets con previsión de RFOA la semana siguiente |

Estructura exacta:

| Rango | Columnas | Campos |
|---|---|---|
| 1 | B:G | IR, SR, Proyecto & Subproyecto, PIVOTE, Proyectos Radio, Real RFOA |
| 2 | I:O | IR, SR, Proyecto & Subproyecto, PIVOTE, Proyectos Radio, Prev RFOA, Real RFOA |
| 3 | Q:W | IR, SR, Proyecto & Subproyecto, PIVOTE, Proyectos Radio, Prev RFOA, Real RFOA |
| 4 | Y:AD | IR, SR, Proyecto & Subproyecto, PIVOTE, Proyectos Radio, Prev RFOA |

`P` y `X` quedan vacías como separadores visuales.

## 8.3 Recuadro de resumen

Rango `AG2:AM16`:

- `AG`: proyecto.
- `AH`: PLANIFICADO Wn-1 — `COUNTIFS(K;proyecto;L;"RFOA")`.
- `AI`: REALIZADO Wn-1.
- `AJ`: REALIZADO NO PLAN Wn-1 — realiza el rango 1.
- `AK`: PLANIFICADO Wn — `COUNTIF(S;proyecto)`.
- `AL`: REALIZADO Wn.
- `AM`: PLANIFICADO Wn+1 — `COUNTIF(AA;proyecto)`.
- Filas 4–15: proyectos; fila 16: `TOTAL` con `SUM`.

Proyectos del resumen: Plan Capacidad - Upgrade, Traslado, JUMPING - Capacidad radio (Portadoras), Plan Capacidad - EVENTOS, Nueva Cobertura - Sin Subproyecto, Cliente, Forced decommissioning, Plan 5G - Commitment, Repliegue, JUMPING - Jumping Radio, JUMPING - Regularizacion Baseline, Plan 5G - NR700 EA.

## 8.4 Otros elementos de PLANIFICACION

- `AH34:AO41`: tabla de KO / re-planificación (SR, JIRA IR, proyecto, causa KO, RE-Planificación).
- `AD48:AD53`: lista auxiliar de 6 IRs.
- Gráfico **“Seguimiento trabajos”** referenciando `AH16`, `AK16`, `AL16`.

## 8.5 Auditoría de macros futuras

### 8.5.1 Macros 7, 8 y 9

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



## 8.6 Implementación (`fase8_adicionales.py`)

Genera un libro **nuevo y limpio**: `resultados/ADICIONALES_SEMI_Limpio.xlsx`,
con solo dos pestañas, sin gráficos ni hojas ocultas. No modifica
`ADICIONALES SEGUIMIENTO.xlsx`.

1. **Foto fija**: si no existe `ADICIONALES SEGUIMIENTO_WN-1.xlsx`, copia el
   original (solo lectura). Si existe, no se toca. El original queda intacto.
2. **DATA IN**: cabeceras de JIRA en fila 1, índice 1…261 en fila 2 y datos
   desde `A3`.
3. **Rango 2 (`I3:O`)**: copia las 7 columnas del Rango 3 (`Q:W`) de la foto
   fija, incluida `Real RFOA` en `O`.
4. **Rango 1 (`B3:G`)**: DATA IN con `Prevision RFOA` (AD) en la semana
   **N-1**, excluyendo los IR ya escritos en Rango 2; última columna
   `Real RFOA` (K).
5. **Rango 3 (`Q3:W`)**: DATA IN con `Prevision RFOA` en la semana **N**
   (Prev RFOA + Real RFOA).
6. **Rango 4 (`Y3:AD`)**: DATA IN con `Prevision RFOA` en la semana **N+1**.
7. **Resumen `AG2:AM16`**: proyectos en AG4:AG15, TOTAL en AG16 y fórmulas
   `COUNTIFS` para:
   - AH Planificado Wn-1
   - AI Realizado Wn-1
   - AJ Realizado no plan Wn-1
   - AK Planificado Wn
   - AL Realizado Wn
   - AM Planificado Wn+1
8. Si hay más de 22 tickets, se escriben filas adicionales por debajo de la
   24. Las fórmulas usan columnas completas, así que siguen contando.

`PIVOTE` siempre es `RFOA`. `P` y `X` quedan vacías como separadores.

El libro se construye con **xlsxwriter**, por lo que el mismo script funciona
en Mac y Windows, sin Excel COM.

### Ejecución

```bat
py automatizacion\fase8_adicionales.py
```

```bash
python3 automatizacion/fase8_adicionales.py
```

Salida: `Nueva carpeta (2)/resultados/ADICIONALES_SEMI_Limpio.xlsx`.

### Dry-run

```bash
python automatizacion/fase8_adicionales.py --dry-run
```

Muestra R1/R2/R3/R4 sin escribir nada.

La app ofrece **📝 Actualizar Adicionales (Fase 8)** y descarga directamente
`ADICIONALES_SEMI_Limpio.xlsx`.
