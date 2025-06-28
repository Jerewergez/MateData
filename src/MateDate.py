import os
import glob
import re
import logging
import time
from datetime import datetime
import win32com.client as win32

# ─── Configuración ─────────────────────────────────────────────────────────────
EXCEL_FILES_DIR = r"C:\ruta\a\los\archivos"
OUTPUT_BASE_DIR = r"C:\ruta\de\salida"
EXCEL_PASSWORD  = "tu_password"

# Patrón: dd-mm justo antes de " al"
DATE_PATTERN = re.compile(r"(\d{2}-\d{2})(?=\s*al)", flags=re.IGNORECASE)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[logging.StreamHandler()]
)


def print_resumen_kpi(wb, stage):
    """
    Recorre cada hoja de wb, busca tablas (ListObjects) cuyo nombre empiece
    con RESUMEN_KPI y las imprime en consola con formato de tabla ASCII.
    stage: 'Before' o 'After'
    """
    for ws in wb.Worksheets:
        try:
            for lo in ws.ListObjects:
                if not lo.Name.upper().startswith("RESUMEN_KPI"):
                    continue

                # Header
                hdr_rng = lo.HeaderRowRange
                cols = hdr_rng.Columns.Count
                headers = []
                for j in range(1, cols + 1):
                    v = hdr_rng.Cells(1, j).Value
                    headers.append(str(v) if v is not None else "")

                # Data
                data_rng = lo.DataBodyRange
                rows_list = []
                if data_rng:
                    for i in range(1, data_rng.Rows.Count + 1):
                        row = []
                        for j in range(1, cols + 1):
                            cell = ws.Cells(
                                data_rng.Row + i - 1,
                                data_rng.Column + j - 1
                            ).Value
                            # Formatear objetos datetime
                            if hasattr(cell, "strftime"):
                                row.append(cell.strftime("%Y-%m-%d"))
                            else:
                                row.append("" if cell is None else str(cell))
                        rows_list.append(row)

                # Calcular anchos de columna
                widths = [len(h) for h in headers]
                for row in rows_list:
                    for idx, val in enumerate(row):
                        widths[idx] = max(widths[idx], len(val))

                # Construir separadores y líneas
                sep = "+" + "+".join("-" * (w + 2) for w in widths) + "+"
                hdr_line = "| " + " | ".join(
                    headers[i].ljust(widths[i]) for i in range(cols)
                ) + " |"

                # Loguear tabla
                logging.info(f"{stage} refresh – Tabla {lo.Name} en hoja {ws.Name}:")
                logging.info(sep)
                logging.info(hdr_line)
                logging.info(sep)
                if rows_list:
                    for row in rows_list:
                        line = "| " + " | ".join(
                            row[i].ljust(widths[i]) for i in range(cols)
                        ) + " |"
                        logging.info(line)
                else:
                    # fila vacía
                    empty = "| " + " | ".join(
                        "".ljust(widths[i]) for i in range(cols)
                    ) + " |"
                    logging.info(empty)
                logging.info(sep)

        except Exception as e:
            logging.debug(f"No se pudo procesar tablas en hoja {ws.Name}: {e}")


def process_excel_files():
    files = glob.glob(os.path.join(EXCEL_FILES_DIR, "*.xlsx")) \
          + glob.glob(os.path.join(EXCEL_FILES_DIR, "*.xlsm"))
    if not files:
        logging.warning("No se encontraron archivos Excel para procesar.")
        return

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.AutomationSecurity = 1  # msoAutomationSecurityForceDisable

    for full_path in files:
        wb = None
        nombre = os.path.basename(full_path)
        logging.info(f"Procesando {nombre}")
        try:
            # 1) Abrir
            wb = excel.Workbooks.Open(
                full_path, False, False, None,
                EXCEL_PASSWORD, EXCEL_PASSWORD
            )

            # Imprimir tabla antes de refrescar
            print_resumen_kpi(wb, "Before")

            # 2) Refresh general
            wb.UpdateLinks = 1
            wb.RefreshAll()
            excel.CalculateUntilAsyncQueriesDone()

            # 3) Refresh pivotes
            for ws in wb.Worksheets:
                for pt in ws.PivotTables():
                    pt.RefreshTable()

            # Imprimir tabla después de refrescar
            print_resumen_kpi(wb, "After")

            # 4) Leer fecha en AA1
            try:
                raw = wb.Worksheets(1).Range("AA1").Value
                if isinstance(raw, datetime):
                    new_date = raw.strftime("%d-%m")
                else:
                    raise ValueError
            except:
                new_date = datetime.now().strftime("%d-%m")
                logging.debug("No se obtuvo fecha de AA1, usando fecha de hoy.")

            # 5) Renombrar y crear carpeta
            base, ext = os.path.splitext(nombre)
            if DATE_PATTERN.search(base):
                nuevo_base = DATE_PATTERN.sub(new_date, base)
            else:
                nuevo_base = f"{base} {new_date}"

            target_dir = os.path.join(OUTPUT_BASE_DIR, nuevo_base)
            os.makedirs(target_dir, exist_ok=True)
            nuevo_path = os.path.join(target_dir, nuevo_base + ext)

            # 6) Guardar copia
            wb.SaveCopyAs(nuevo_path)
            logging.info(f"Guardado → {nuevo_path}")

        except Exception as e:
            logging.error(f"Error procesando {nombre}: {e}")
        finally:
            if wb:
                wb.Close(False)

    excel.Quit() 


if __name__ == "__main__":
    start = time.time()
    logging.info("==== Se inicio 🧉 MateDate correctamente ====")
    process_excel_files()
    logging.info(f"==== FIN EN {time.time() - start:.2f}s ====")
