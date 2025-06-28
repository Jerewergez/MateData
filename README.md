# 📊 Excel Refresher & Archiver

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

Instalación de dependencias:

```
pip install pywin32
