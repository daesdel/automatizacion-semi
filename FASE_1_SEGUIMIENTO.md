# Fase 1 — Seguimiento E2E

> [← Volver al índice general](DOCUMENTACION_PROYECTO.md)

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

