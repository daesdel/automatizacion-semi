# Fase 6 — Auditoría de tickets

> [← Volver al índice general](DOCUMENTACION_PROYECTO.md)

`fase6_auditoria.py` compara con la columna A de `PREVISIONES JUAN` solo los
**IR de radio** (`IR-...`) que estén en `drive` **o** tengan `Fecha PO > 0` en
`DATA IN`. Excluye DR, RE y tickets sin actividad. Genera
`Auditoria_Tickets_Faltantes.xlsx` con IR, Origen, H, AB y AJ.

La app usa el **`JIRA.xlsx` generado por la Fase 1** en la carpeta de salida
(no `JIRA con DR.xlsx`). Si no existe o no contiene `IR_VDF`, muestra un error
amigable indicando que se ejecute primero el proceso principal.

El botón **🔎 Ejecutar Auditoría de Tickets** tiene la misma experiencia que el
principal: cronómetro con fecha/hora, barra de progreso 0–100 %, consola en
vivo, mensaje de éxito y descarga.

