# Fase 7 — Fichero de trabajo

> [← Volver al índice general](DOCUMENTACION_PROYECTO.md)

`fase7_fichero.py` genera `resultados/DDMMYYYY_FICHERO.xlsx` con:

- **PREVISIONES JUAN**: copia la hoja completa. Si la Fase 5 pudo recalcular
  con Excel, copia valores; si no, copia las fórmulas y conserva `DATA IN`
  oculta para que Excel las calcule al abrir.
- **RFOA SIN PREVISION**: lee `Previsiones de RFOA_central.xlsx` → `Nodos`,
  tabla `Tabla1`, y copia solo las filas visibles (si hay filtros), con tabla
  `tbl_RFOA_SIN_PREVISION`.
- **DUPLICADOS**: detecta IRs repetidos en PREVISIONES (columna A) y filas
  idénticas en RFOA SIN PREVISION; pinta de morado y lista TICKET/HOJA.
- Los encabezados de RFOA SIN PREVISION se fuerzan a **Calibri 11 negrita,
  color negro RGB `FF000000` y fondo blanco** (sin estilo de tabla que los
  vuelva a pintar de blanco).

Antes de abrir, `fase7_fichero.py` valida que PREVISIONES y el RFOA central son
ZIP/XLSX válidos, con reintentos si el archivo se está escribiendo. Si están
dañados o abiertos en Excel, muestra un error amigable y no un Traceback.

La app incorpora **📁 Generar Fichero (Fase 7)** con cronómetro, barra de
progreso, consola en vivo y descarga del FICHERO.

La hoja **RFOA SIN PREVISION** clona estilos, anchos, altos y colores de las
celdas de `Nodos` (tabla `Tabla1`), además de crear
`tbl_RFOA_SIN_PREVISION`.



> El cálculo de valores y el uso de Excel COM/pycel se detalla en [Fase 5](FASE_5_PREVISIONES.md#610-calculo-excel-com-o-pycel).
