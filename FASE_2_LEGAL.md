# Fase 2 — Legal (Ministerio VDF/OSP)

> [← Volver al índice general](DOCUMENTACION_PROYECTO.md)

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

