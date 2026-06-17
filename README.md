# ExelenParalelo

Procesamiento **paralelo** de archivos Excel con PowerQuery, ODBC y tablas dinámicas.
Abre cada workbook, refresca todas las conexiones, espera a que terminen, actualiza
PivotTables y guarda una copia renombrada según la celda **AA1**.

---

## Tabla de contenidos

- [Arquitectura](#arquitectura)
- [Requisitos](#requisitos)
- [Instalación](#instalación)
- [Configuración](#configuración)
- [Cómo funciona](#cómo-funciona)
- [Flujo por worker](#flujo-por-worker)
- [Salida y logs](#salida-y-logs)
  - [Terminal (stdout)](#terminal-stdout)
  - [Resumen final](#resumen-final)
  - [Tabla comparativa acumulada](#tabla-comparativa-acumulada)
- [CSV de auditoría](#csv-de-auditoría)
- [Manejo de errores](#manejo-de-errores)
- [Preguntas frecuentes](#preguntas-frecuentes)

---

## Arquitectura

```
┌──────────────────────────────────────────────────────┐
│                    PROCESO PADRE                      │
│  ┌──────────┐   ┌──────────┐   ┌──────────┐         │
│  │ Worker 1 │   │ Worker 2 │   │ Worker 3 │  … hasta │
│  │ (Excel)  │   │ (Excel)  │   │ (Excel)  │  MAX_   │
│  └────┬─────┘   └────┬─────┘   └────┬─────┘  WORKERS│
│       │              │              │               │
│       └──────────────┴──────────────┘               │
│                         │ IPC: multiprocessing.Queue │
│                         ▼                            │
│              ┌────────────────────┐                  │
│              │   Orquestador      │                  │
│              │  - timeouts        │                  │
│              │  - reintentos      │                  │
│              │  - barra progreso  │                  │
│              │  - CSV auditoría   │                  │
│              └────────────────────┘                  │
└──────────────────────────────────────────────────────┘
```

Cada worker corre en un **proceso hijo independiente** (`multiprocessing.Process`).
Se comunican con el padre mediante una `Queue` (por mensajes): primero envían el
PID de su instancia de Excel, y al terminar envían el resultado (éxito/error,
ruta del archivo guardado, stdout del proceso).

El padre mantiene un **pool de hasta `MAX_WORKERS` workers concurrentes**.
apenas uno termina, lanza el siguiente de la cola de pendientes.

---

## Requisitos

| Requisito | Detalle |
|-----------|---------|
| **SO** | Windows (usa `win32com`, `win32process`, `taskkill`) |
| **Excel** | Microsoft Excel instalado (2016+) |
| **Python** | 3.8 o superior |
| **pywin32** | `pip install pywin32` |
| **colorama** | Opcional — `pip install colorama` (logs en color) |

No funciona en Linux/Mac por las dependencias COM de Windows.

---

## Instalación

```powershell
# 1. Clonar o copiar el script
# 2. Instalar dependencias
pip install pywin32 colorama

# 3. Verificar que funcione
python ExelenParalelo.py --help   # (o simplemente ejecutarlo)
```

> **Nota**: `colorama` es opcional. Sin ella los logs se ven igual, solo sin colores.

---

## Configuración

Toda la configuración está al inicio del archivo, en la sección `CONFIG`.
Son variables globales con valores por defecto que cubren el caso típico.

| Variable | Tipo | Defecto | Qué hace |
|----------|------|---------|----------|
| `INPUT_EXCEL_DIR` | `str` | `D:\Bases` | Carpeta con los excels a procesar |
| `LOCAL_OUTPUT_DIR` | `Path` | `D:\Procesados` | Carpeta donde se guardan las copias procesadas |
| `EXCEL_PASSWORD` | `str` | `"vam123"` | Contraseña de apertura de los workbooks |
| `CLOSE_ALL_EXCEL_BEFORE` | `bool` | `True` | Mata toda instancia de Excel **antes** de empezar |
| `CLOSE_ALL_EXCEL_AFTER` | `bool` | `True` | Mata toda instancia de Excel **al terminar** |
| `CLOSE_WAIT_S` | `int` | `3` | Segundos de espera después de matar Excel |
| `DATE_PATTERN` | `re.Pattern` | `(\d{2}-\d{2})(?=\s*al)` | Regex para detectar fechas en el nombre del archivo |
| `MAX_WORKERS` | `int` | `3` | Máximo de procesos Excel concurrentes |
| `TIMEOUT_S` | `int` | `3600` | Timeout por archivo (1 hora) |
| `MAX_RETRIES` | `int` | `2` | Reintentos antes de dar un archivo por perdido |

### Cómo modificar

Abrís el archivo, cambiás la variable y guardás. No necesita archivos
de configuración externos ni variables de entorno.

```python
# Ejemplo: aumentar workers a 5 y reducir timeout a 30 min
MAX_WORKERS = 5
TIMEOUT_S   = 1800
```

---

## Cómo funciona

### 1. Inicio

1. Si `CLOSE_ALL_EXCEL_BEFORE` está activo, mata todas las instancias de
   `EXCEL.EXE` con `taskkill /F`. Esto evita conflictos de archivos bloqueados.
2. Escanea `INPUT_EXCEL_DIR` en busca de archivos `*.xls*`.
3. Crea el CSV de auditoría con timestamp en el nombre.

### 2. Procesamiento paralelo

El script mantiene hasta `MAX_WORKERS` workers vivos simultáneamente:

```
Cola de pendientes: [A.xlsx, B.xlsx, C.xlsx, D.xlsx, E.xlsx]
                          │
             ┌────────────┼────────────┐
             ▼            ▼            ▼
         Worker A     Worker B     Worker C
         (activo)     (activo)     (activo)
             │
             ▼  (termina)
         [A.xlsx OK]
             │
             ▼
         Lanzar D.xlsx → reemplaza a A en el pool
```

### 3. Finalización

Cuando todos los archivos se procesaron (o agotaron reintentos), se muestra
un **resumen final** con la cuenta de éxitos y errores.

---

## Flujo por worker

Cada worker corre el siguiente pipeline:

```
 1. Iniciar Excel invisible
    └─ win32.DispatchEx("Excel.Application")
    └─ Enviar PID al padre (para poder matarlo si hay timeout)

 2. Abrir workbook
    └─ excel.Workbooks.Open(..., password)

 3. Snapshot "Before" de tablas RESUMEN_KPI
    └─ (solo si existen)

 4. RefreshAll
    └─ disable_background_refresh() → apaga BackgroundQuery
    └─ wb.RefreshAll()
    └─ wait_until_queries_done() → polling hasta que todas
       las conexiones terminen (con timeout de 30 min)

 5. Refrescar PivotTables
    └─ Itera todas las hojas → todas las tablas dinámicas

 6. Snapshot "After" de RESUMEN_KPI

 7. Guardar copia local
    └─ Lee celda AA1 → extrae fecha
    └─ Renombra: reemplaza "dd-mm" en el nombre original por la fecha de AA1
    └─ wb.SaveCopyAs()

 8. Cerrar Excel

 9. Enviar resumen al padre
```

---

## Salida y logs

El script produce tres tipos de salida:

### Terminal (stdout)

Logs en vivo con timestamps y colores (si colorama está instalado):

```
  ┌────────────────────────────────────────────┐
  │  INICIO — ExelenParalelo                   │
  └────────────────────────────────────────────┘

  [14:30:01] ✓ Archivos detectados: 5

  [14:30:01] ▶ Ventas_Q1 (PID 8412, intento 1/2)
  📊 ███░░░░░░░░░░░░░░░░░ 1/5  (3 activos, 1 pendientes)  12s  (20.0%)

  [14:30:12] ✓ Ventas_Q1 → D:\Procesados\Ventas 15-03.xlsx (11.2s)
  [14:30:12] ▶ Stock_Actual (PID 8501, intento 1/2)
  📊 ██████░░░░░░░░░░░░░░ 2/5  (3 activos, 0 pendientes)  14s  (40.0%)
```

Cada worker vuelca su stdout interno indentado con `┊` para distinguirlo
de los logs del padre.

### CSV de auditoría

Se genera un archivo `Datos_De_Ejecucion_<timestamp>.csv` en `LOCAL_OUTPUT_DIR`
con una fila por cada intento de procesamiento (incluyendo reintentos).

| Columna | Descripción |
|---------|-------------|
| `timestamp` | Momento del log |
| `intento` | Número de intento (1, 2, …) |
| `archivo` | Nombre del archivo original |
| `ok` | `True` o `False` |
| `ruta_local` | Ruta absoluta del archivo guardado |
| `fecha_AA1` | Fecha extraída de la celda AA1 |
| `duracion_s` | Segundos que tomó procesarlo |
| `error` | Mensaje de error (si falló) |

Las columnas están separadas por `;` (punto y coma) para que Excel las abra
directamente.

### Resumen final

Al terminar todo, se imprime un bloque de resumen:

```
  ┌────────────────────────────────────────────┐
  │  RESUMEN FINAL                             │
  └────────────────────────────────────────────┘

  [14:35:40] ✓ Procesados: 5
  [14:35:40] ✓ Éxitos:     4
  [14:35:40] ✗ Errores:    1
               • Balance_Anual: Timeout 3600s

  [14:35:40] Duración total: 320.45s
  [14:35:40] CSV: D:\Procesados\Datos_De_Ejecucion_20250617_143001.csv
```

### Tabla comparativa acumulada

Después del resumen final, el script muestra una **tabla consolidada** con todos los KPIs
de la tabla `RESUMEN_KPI` de **todos los archivos procesados**, comparando los valores
Before vs After del refresh. Ideal para ver de un vistazo qué cambió en cada indicador.

```
  ┌────────────────────────────────────────────────────────────────────────────────┐
  │  Tabla consolidada — todas las tablas RESUMEN_KPI ordenadas por SERVICIO      │
  └────────────────────────────────────────────────────────────────────────────────┘

  ARCHIVO      │ KPI        │ SUBÁREA │ SERVICIO     │ ANTES   │ DESPUÉS │ Δ GAP    │ ALCANCE
  ───────────────────────────────────────────────────────────────────────────────────────
  Ventas.xlsx  │ GxH        │ SOPORTE │ SOPORTE-RRSS │ 3.7383  │ 3.8157  │ +0.0774  │ Pasa
  Ventas.xlsx  │ REL        │ SOPORTE │ SOPORTE-RRSS │ 0.7496  │ 0.7478  │ -0.0019  │ Pasa
 ⚡ Ventas.xlsx │ % CASO     │ SOPORTE │ SOPORTE-RRSS │ 0.7875  │ 0.7793  │ -0.0081  │ NO Pasa
  ───────────────────────────────────────────────────────────────────────────────────────
  Stock.xlsx   │ GxH        │ CALIDAD │ CALIDAD-SAT   │ 2.4500  │ 2.5100  │ +0.0600  │ Pasa
```

| Característica | Detalle |
|----------------|---------|
| **Origen** | Datos extraídos de las tablas `RESUMEN_KPI` de cada Excel |
| **Orden** | Por SERVICIO (3ra columna original) |
| **Archivo** | Columna adicional para identificar el origen cuando hay múltiples archivos |
| **Δ GAP** | `DESPUÉS - ANTES` del RESULTADO (col 4). En **verde** si mejoró, **rojo** si empeoró |
| **ALCANCE** | Col 5 (PASA/NO PASA). Se muestra en **rojo** si cambió entre Before y After |
| **⚡** | Marcador en filas donde se detectó algún cambio (gap ≠ 0 o ALCANCE cambió) |

**Nota:** Si `colorama` no está instalado, los colores no se ven pero la tabla se muestra igual.

---

## CSV de auditoría

Cada ejecución genera un CSV en `LOCAL_OUTPUT_DIR` con timestamp en el nombre:

```
D:\Procesados\
├── Ventas 15-03.xlsx
├── Stock 22-03.xlsx
├── Datos_De_Ejecucion_20250617_143001.csv   ← este
└── Datos_De_Ejecucion_20250618_090215.csv   ← ejecución anterior
```

El CSV es un archivo de texto plano con encoding UTF-8 y separador `;`.
Podés abrirlo directamente en Excel o PowerQuery para hacer análisis.

---

## Manejo de errores

### Timeout por archivo (`TIMEOUT_S`)

Si un worker supera el tiempo límite:

1. Se **termina** el proceso worker (`p.terminate()`)
2. Se **mata** la instancia de Excel asociada (vía `taskkill /PID`)
3. Si quedan reintentos, se **reencola** el archivo
4. Si se agotaron los reintentos, se marca como error definitivo

### Error interno del worker

Cualquier excepción dentro del worker:

1. Se cierran los objetos COM (workbook, Excel) con `try/except`
2. Se envía el resumen con `ok=False` y el mensaje de error
3. El padre decide si reintenta según `MAX_RETRIES`

### Trabajo sucio (cleanup)

- Si se cae el padre (Ctrl+C, cierre de terminal), los workers hijos
  se vuelven huérfanos.
- `CLOSE_ALL_EXCEL_AFTER=True` mata todo Excel al finalizar (incluso
  instancias pre-existentes), actuando como red de seguridad.
- Si esto es un problema, ponerlo en `False`.

---

## Preguntas frecuentes

### ¿Puedo cambiar la cantidad de workers sobre la marcha?

No — hay que editar `MAX_WORKERS` en el archivo y reiniciar.

### ¿Qué pasa si un Excel está abierto por un usuario?

Si `CLOSE_ALL_EXCEL_BEFORE=True`, se cierra forzosamente (se pierden
cambios no guardados). Si está en `False`, `Workbooks.Open()` va a fallar
porque el archivo está bloqueado.

### ¿Cómo sé que archivos fallaron?

1. Mirá el **resumen final** en terminal
2. Abrí el **CSV de auditoría** y filtrá por `ok=False`

### ¿Puedo procesar solo algunos archivos?

Mové los que no querés procesar fuera de `INPUT_EXCEL_DIR`, o cambiá
la variable para que apunte a una subcarpeta.

### ¿Sirve en Linux con LibreOffice?

No. El script depende de `win32com` (COM de Windows) y Excel.
Para Linux necesitarías una aproximación completamente distinta
(por ejemplo, `python-calc` con LibreOffice o `xlwings` vía REST).

---

## Licencia

Uso interno — Pulso Data Team.
