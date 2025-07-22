# 🧉 MateDate 📊 Excel Refresher & Archiver

Este script automatiza la apertura, actualización y resguardo de archivos de Excel con Power Query y Tablas Dinámicas. Ideal para flujos de trabajo que requieren actualización periódica de reportes con mínima intervención manual.

## 🚀 Funcionalidades

- 🔄 Actualiza consultas de Power Query y tablas dinámicas.
- 🧾 Muestra en consola tablas llamadas `RESUMEN_KPI` antes y después del refresh.
- 📆 Lee la fecha de la hoja resumen desde la celda `AA1` para renombrar copias procesadas.
- 📂 Guarda versiones organizadas por fecha en carpetas específicas.

## 🛠️ Requisitos

- Sistema operativo: **Windows**
- Software necesario: **Microsoft Excel**
- Python 3.7 o superior
- Librerías necesarias:
  - `pywin32`
  - `os`, `glob`, `re`, `logging`, `time`, `datetime` (incluidas en Python)

Instalación de dependencias:

```
pip install pywin32
```
## ⚙️ Configuración

Editá las siguientes variables al inicio del script:

```python
EXCEL_FILES_DIR = r"C:\ruta\a\los\archivos"
OUTPUT_BASE_DIR = r"C:\ruta\de\salida"
EXCEL_PASSWORD  = "tu_password"
```

## 🧪 Ejecución
Desde la terminal, ejecutá:
```
python actualizar_excel.py
```

Se logueará el progreso, incluyendo un resumen ASCII de las tablas RESUMEN_KPI:
```
2025-06-28 14:31:28 - INFO - ==== Se inicio la actualización correctamente ====
2025-06-28 14:31:30 - INFO - Procesando CxH SERVICE RRSS SOPORTE al.xlsx
2025-06-28 14:31:36 - INFO - Before refresh – Tabla RESUMEN_KPI1 en hoja RESUMEN:
2025-06-28 14:31:36 - INFO - +----------------+---------+--------------+--------------------+---------+------------+
2025-06-28 14:31:36 - INFO - | KPI            | SUBÁREA | SERVICIO     | RESULTADO          | ALCANCE | FECHA KPI  |
2025-06-28 14:31:36 - INFO - +----------------+---------+--------------+--------------------+---------+------------+
2025-06-28 14:31:36 - INFO - | GxH            | SOPORTE | SOPORTE-RRSS | 3.7383450564344636 | 0.01    | 2025-06-21 |
2025-06-28 14:31:36 - INFO - | REL            | SOPORTE | SOPORTE-RRSS | 0.7496116701560073 | 0.0     | 2025-06-21 |
2025-06-28 14:31:36 - INFO - | % CASO CERRADO | SOPORTE | SOPORTE-RRSS | 0.7874595231086252 | Pasa    | 2025-06-21 |
2025-06-28 14:31:36 - INFO - +----------------+---------+--------------+--------------------+---------+------------+
2025-06-28 14:32:14 - INFO - After refresh – Tabla RESUMEN_KPI1 en hoja RESUMEN:
2025-06-28 14:32:14 - INFO - +----------------+---------+--------------+--------------------+---------+------------+
2025-06-28 14:32:14 - INFO - | KPI            | SUBÁREA | SERVICIO     | RESULTADO          | ALCANCE | FECHA KPI  |
2025-06-28 14:32:14 - INFO - +----------------+---------+--------------+--------------------+---------+------------+
2025-06-28 14:32:14 - INFO - | GxH            | SOPORTE | SOPORTE-RRSS | 3.8156587308734617 | 0.02    | 2025-06-25 |
2025-06-28 14:32:14 - INFO - | REL            | SOPORTE | SOPORTE-RRSS | 0.7477529164276152 | 0.0     | 2025-06-25 |
2025-06-28 14:32:14 - INFO - | % CASO CERRADO | SOPORTE | SOPORTE-RRSS | 0.7793468200449557 | Pasa    | 2025-06-25 |
2025-06-28 14:32:14 - INFO - +----------------+---------+--------------+--------------------+---------+------------+
2025-06-28 14:32:18 - INFO - Guardado → C:\Users\i44475827\Documents\Automatizaciones\Procesados\CxH SERVICE RRSS SOPORTE al 25-06\CxH SERVICE RRSS SOPORTE al 25-06.xlsx
```

## 🗂️ Estructura esperada
```
Automatizaciones/
├── archivo_original.xlsx
├── Procesados/
│   └── archivo_original 28-06/
│       └── archivo_original 28-06.xlsx
```
## 🔐 Seguridad

- El script ejecuta Excel en segundo plano y desactiva avisos.
- La contraseña está escrita en el código. Podés moverla a una variable de entorno para mayor seguridad.

