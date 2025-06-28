# 📊 Excel MateDate 🧉

Este script automatiza la apertura, actualización y resguardo de archivos de Excel con Power Query y Tablas Dinámicas. Ideal para flujos de trabajo que requieren actualización periódica de reportes con mínima intervención manual.

## 🚀 Funcionalidades

- 🔄 Actualiza consultas de Power Query y tablas dinámicas.
- 🧾 Muestra en consola tablas llamadas `RESUMEN_KPI` antes y después del refresh.
- 📆 Lee la fecha desde la celda `AA1` para renombrar copias procesadas.
- 📂 Guarda versiones organizadas por fecha en carpetas específicas.

## 🛠️ Requisitos

- Sistema operativo: **Windows**
- Software necesario: **Microsoft Excel**
- Python 3.7 o superior
- Librerías necesarias:
  - `pywin32`
  - `os`, `glob`, `re`, `logging`, `time`, `datetime` (incluidas en Python)

⚙️ Configuración
Editá las siguientes variables al inicio del script:

python
EXCEL_FILES_DIR = r"C:\ruta\a\los\archivos"
OUTPUT_BASE_DIR = r"C:\ruta\de\salida"
EXCEL_PASSWORD  = "tu_password"
🧪 Ejecución
Desde la terminal, ejecutá:

bash
python actualizar_excel.py
Se logueará el progreso, incluyendo un resumen ASCII de las tablas RESUMEN_KPI.

🗂️ Estructura esperada
Automatizaciones/
├── archivo_original.xlsx
├── Procesados/
│   └── archivo_original 28-06/
│       └── archivo_original 28-06.xlsx
🔐 Seguridad
El script ejecuta Excel en segundo plano y desactiva avisos.

La contraseña está escrita en el código. Podés moverla a una variable de entorno para mayor seguridad.
