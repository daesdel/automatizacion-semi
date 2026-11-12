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
| 2 | `I1:N24` | `SEMANA N-1` | Tickets que la semana anterior tenían previsión de RFOA en esa semana |
| 3 | `Q1:W24` | `SEMANA N` | Tickets que la semana anterior tenían previsión de RFOA en la semana actual |
| 4 | `Y1:AD24` | `SEMANA N+1` | Tickets con previsión de RFOA la semana siguiente |

Columnas de cada bloque: JIRA IR, SR, PROYECTO & SUBPROYECTO, PIVOTE, Proyectos Radio y fechas (Real RFOA / Prev RFOA según bloque).

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

1. **Foto fija**: si no existe `ADICIONALES SEGUIMIENTO_WN-1.xlsx`, copia el
   libro actual. Si existe, no se toca.
2. **DATA IN**: vuelca `JIRA.xlsx` (`IR_VDF`, desde la fila 5) en
   `DATA IN!A3`, limpiando antes las filas antiguas.
3. **Rango 2 (`I3:N24`)**: copia las 6 primeras columnas del Rango 3 de la
   foto fija (`Q:V`). La columna `W` (Real RFOA) no cabe en `I:N`; si se
   necesitara, habría que ampliar el rango a `O`.
4. **Rango 1 (`B3:G24`)**: tickets del DATA IN con `Prevision RFOA` (AD) en la
   semana **N-1**, excluyendo los IR ya escritos en el Rango 2. La última
   columna es `Real RFOA` (K).
5. **Rango 3 (`Q3:W24`)**: tickets con `Prevision RFOA` en la semana **N**;
   se rellenan también Prev RFOA y Real RFOA.
6. **Rango 4 (`Y3:AD24`)**: tickets con `Prevision RFOA` en la semana **N+1**.
7. **Resumen**: las fórmulas `COUNTIFS`/`SUMPRODUCT` de `AG2:AM16` se
   mantienen; usan columnas completas o rangos hasta la fila 1998, por lo que
   siguen contando si se añaden filas.
8. Si hay más de 22 tickets en un rango, se insertan filas a partir de la 25
   y se desplazan KO/re-planificación y listas auxiliares.

La columna PIVOTE se escribe siempre como `RFOA`.

### Ejecución en Windows (producción)

```bat
py automatizacion\fase8_adicionales.py
```

Usa Excel COM: actualiza DATA IN y Planificación conservando gráficos.

### Prueba en Mac

Mac no tiene Excel COM; se usa openpyxl, que puede tardar y **perder
gráficos**. No se debe ejecutar sobre el archivo original. Prueba segura:

```bash
cp "Nueva carpeta (2)/ADICIONALES SEGUIMIENTO.xlsx" /tmp/ADICIONALES_TEST.xlsx
python automatizacion/fase8_adicionales.py \
  --adicionales=/tmp/ADICIONALES_TEST.xlsx \
  --backup=/tmp/ADICIONALES_TEST_WN-1.xlsx \
  --jira="Nueva carpeta (2)/JIRA.xlsx" \
  --dry-run
```

El `--dry-run` calcula y muestra R1/R2/R3/R4 sin escribir. Para probar la
escritura, quita `--dry-run` y revisa las hojas del archivo de `/tmp`.
