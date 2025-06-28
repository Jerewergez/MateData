# MateData 🧉
Este script automatiza la apertura, actualización y resguardo de archivos de Excel con Power Query y Tablas Dinámicas. Ideal para flujos de trabajo que requieren actualización periódica de reportes con mínima intervención manual.

🚀 Funcionalidades
Abre automáticamente todos los archivos .xlsx y .xlsm de una carpeta.

Actualiza consultas de Power Query y tablas dinámicas.

Muestra en consola tablas llamadas RESUMEN_KPI antes y después del refresh.

Lee la fecha desde la celda AA1 para nombrar las copias procesadas.

Renombra los archivos y los guarda organizadamente en una carpeta destino con fecha.

🛠️ Requisitos
Windows con Microsoft Excel instalado.

Python 3.7 o superior.

Dependencias:

pywin32 (automatización de Excel)

Librerías estándar: os, glob, re, logging, time, datetime
