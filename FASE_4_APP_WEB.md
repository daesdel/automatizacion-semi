# Fase 4 — Interfaz Streamlit

> [← Volver al índice general](DOCUMENTACION_PROYECTO.md)

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
- **Sanity check en segundo plano con python-calamine**: un proceso
  independiente (`sanity_tickets.py`) cuenta JIRA/Hoja 2 en ~8 s y compara con
  el DATA IN anterior.
  La UI muestra barra, consola y botón de parada; si baja, aparece el aviso con
  **Sí, continuar / Cancelar**. Nunca bloquea la interfaz. Si `python-calamine`
  no está instalado, la lectura vuelve a openpyxl y se vuelve lenta.
- Botón **🛑 Detener Proceso** mientras hay una ejecución: mata el proceso hijo
  y deja la app operativa.
- **Registro persistente** en `st.session_state`: Pipeline, Auditoría, Sanity
  y Fase 7 se acumulan con hora y se muestran en pestañas dentro de un
  expander. El título del pipeline es **“Pipeline (Fases 1, 2 y 5)”**.
- **Cronómetro** con fecha/hora de inicio y tiempo transcurrido.
- **Barra de progreso real** por fases (1, 2 y 5).
- **Consola en vivo** con las últimas líneas de `print`.
- Botón **🔎 Ejecutar Auditoría de Tickets** (Fase 6).
- Botón **📁 Generar Fichero (Fase 7)**, con campo opcional para la ruta de
  `Previsiones de RFOA_central.xlsx`.
- Fases 6 y 7 se pueden lanzar **sin pipeline previo**: buscan
  `Seguimiento_Previsiones_Limpio.xlsx` en `resultados_app`, en la carpeta de
  datos, en `resultados` o en `subidas/resultados_app`.
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

