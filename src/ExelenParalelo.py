#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ExelenParalelo.py — Procesamiento paralelo de archivos Excel con PowerQuery

Procesa archivos Excel (.xlsx/.xlsb) en paralelo: abre cada workbook,
refresca todas las conexiones (PowerQuery, ODBC, OLEDB), actualiza tablas
dinámicas y guarda una copia renombrada según la celda AA1.

Arquitectura
────────────
  Worker pool (multiprocessing.Process) con comunicación vía Queue.
  Cada worker lanza una instancia invisible de Excel via win32com.
  El padre orquesta hasta N workers concurrentes con timeouts y reintentos.

Requisitos
──────────
  - Windows (depende de win32com)
  - Microsoft Excel instalado
  - Python 3.8+
  - pywin32
  - colorama (opcional — logs en color)

Uso
───
  python ExelenParalelo.py
"""

import os
import sys
import re
import glob
import time
import csv
import queue
import subprocess
from datetime import datetime
from io import StringIO
from pathlib import Path
from multiprocessing import Process, Queue
from typing import Dict, Tuple, List, Optional

import win32com.client as win32
import win32process

# ─────────────────────────────── LOGGING MEJORADO ───────────────────────────────

try:
    import colorama
    colorama.init()
    _HAVE_COLOR = True
except ImportError:
    _HAVE_COLOR = False

# Pre-compute ANSI codes para evitar backslashes en f-strings
_C_CYAN  = "\033[96m"
_C_GREEN = "\033[92m"
_C_YELLOW = "\033[93m"
_C_RED   = "\033[91m"
_C_GRAY  = "\033[90m"
_C_RESET = "\033[0m"


def _c(code: str, text: str) -> str:
    """Wrap text in ANSI color code if colorama is available."""
    return f"{code}{text}{_C_RESET}" if _HAVE_COLOR else text


def _ts() -> str:
    return datetime.now().strftime("%H:%M:%S")


def info(msg: str):
    print(f"  [{_ts()}] {msg}")


def ok(msg: str):
    print(f"  [{_ts()}] {_c(_C_GREEN, '✓')} {msg}")


def warn(msg: str):
    print(f"  [{_ts()}] {_c(_C_YELLOW, '⚠')} {msg}")


def error(msg: str):
    print(f"  [{_ts()}] {_c(_C_RED, '✗')} {msg}")


def section(title: str):
    cols = getattr(os, 'get_terminal_size', lambda: (80, 24))().columns - 4
    sep = "─" * min(60, max(20, cols))
    top = f"  {_c(_C_CYAN, '┌' + sep + '┐')}"
    mid = f"  {_c(_C_CYAN, '│')}  {title}"
    bot = f"  {_c(_C_CYAN, '└' + sep + '┘')}"
    print(f"\n{top}\n{mid}\n{bot}\n")


# ───────────────────────────────────── CONFIG ─────────────────────────────────────
INPUT_EXCEL_DIR   = r"D:\Bases"          # Excels a procesar (con conexiones ya correctas)
LOCAL_OUTPUT_DIR  = Path(r"D:\Procesados")

EXCEL_PASSWORD      = "vam123"
CLOSE_ALL_EXCEL_BEFORE = True
CLOSE_ALL_EXCEL_AFTER  = True
CLOSE_WAIT_S = 3  # segundos

DATE_PATTERN = re.compile(r"(\d{2}-\d{2})(?=\s*al)", flags=re.IGNORECASE)

MAX_WORKERS     = 3
TIMEOUT_S       = 3600     # 1 hora por archivo
MAX_RETRIES     = 2

RUN_TS             = datetime.now().strftime("%Y%m%d_%H%M%S")
CSV_PROCESS_PATH   = LOCAL_OUTPUT_DIR / f"Datos_De_Ejecucion_{RUN_TS}.csv"
CSV_PROCESS_HEADER = [
    "timestamp", "intento", "archivo", "ok",
    "ruta_local", "fecha_AA1", "duracion_s", "error"
]

# ─────────────────────────── HELPERS SISTEMA / CSV ───────────────────────────
def close_all_excels_before_start(force: bool = True):
    """Cierra todas las instancias de EXCEL.EXE (pierde cambios no guardados)."""
    try:
        flag = "/F" if force else ""
        cmd = ["taskkill", flag, "/IM", "EXCEL.EXE", "/T"]
        cmd = [c for c in cmd if c]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
        if res.returncode in (0, 128):
            info("EXCEL.EXE cerrado (si existía).")
        else:
            warn(f"taskkill code {res.returncode}: {res.stdout.strip()} {res.stderr.strip()}")
    except Exception as e:
        warn(f"No se pudo cerrar EXCEL.EXE: {e}")


def append_csv_header_if_missing(csv_path: Path, header: List[str]):
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    if not csv_path.exists():
        with open(csv_path, "w", encoding="utf-8", newline="") as cf:
            csv.writer(cf, delimiter=';').writerow(header)


def append_csv_row(csv_path: Path, row: List[str]):
    with open(csv_path, "a", encoding="utf-8", newline="") as cf:
        csv.writer(cf, delimiter=';').writerow(row)


# ─────────────────────────── REFRESH ROBUSTO ───────────────────────────
def disable_background_refresh(wb):
    """Apaga BackgroundQuery en QueryTables y conexiones ODBC/OLEDB
    para forzar refresh sincrónico."""
    try:
        for ws in wb.Worksheets:
            try:
                for lo in ws.ListObjects:
                    try:
                        qt = lo.QueryTable
                        if qt is not None:
                            qt.BackgroundQuery = False
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass
    try:
        for cn in wb.Connections:
            try:
                o = getattr(cn, "ODBCConnection", None)
                if o is not None:
                    o.BackgroundQuery = False
                o = getattr(cn, "OLEDBConnection", None)
                if o is not None:
                    o.BackgroundQuery = False
            except Exception:
                pass
    except Exception:
        pass


def wait_until_queries_done(excel, wb, timeout_s=1800, poll_s=0.5):
    """
    Espera a que terminen refreshes/consultas.

    1) Intenta Application.CalculateUntilAsyncQueriesDone()
    2) Si falla, hace polling de conexiones/QueryTables/CalculationState.
    """
    start = time.time()

    # 1) Intento directo
    try:
        if hasattr(excel, "CalculateUntilAsyncQueriesDone"):
            excel.CalculateUntilAsyncQueriesDone()
            time.sleep(0.3)
            return True
    except Exception:
        pass

    # 2) Plan B: polling
    def any_refreshing():
        try:
            for cn in wb.Connections:
                o = getattr(cn, "ODBCConnection", None)
                if o is not None and getattr(o, "Refreshing", False):
                    return True
                o = getattr(cn, "OLEDBConnection", None)
                if o is not None and getattr(o, "Refreshing", False):
                    return True
        except Exception:
            pass
        try:
            for ws in wb.Worksheets:
                for lo in ws.ListObjects:
                    qt = getattr(lo, "QueryTable", None)
                    if qt is not None and getattr(qt, "Refreshing", False):
                        return True
        except Exception:
            pass
        try:
            state = getattr(excel, "CalculationState", 0)
            # 0=xlDone, 1=xlCalculating, 2=xlPending
            if state in (1, 2):
                return True
        except Exception:
            pass
        return False

    while any_refreshing():
        if (time.time() - start) > timeout_s:
            raise TimeoutError(f"Espera de refresh superó {timeout_s}s")
        time.sleep(poll_s)

    time.sleep(0.2)
    return True


# ─────────────────────────── LÓGICA DE EXCEL / ETL ───────────────────────────
def print_resumen_kpi(wb, stage, output):
    """Vuelca tablas RESUMEN_KPI a 'output' con formato de grilla."""
    for ws in wb.Worksheets:
        try:
            for lo in ws.ListObjects:
                if not lo.Name.upper().startswith("RESUMEN_KPI"):
                    continue
                hdr_rng = lo.HeaderRowRange
                cols = hdr_rng.Columns.Count
                headers = [str(hdr_rng.Cells(1, j).Value or "") for j in range(1, cols + 1)]
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

                sep_line = "+" + "+".join("-" * (w + 2) for w in widths) + "+"
                hdr_line = "| " + " | ".join(
                    headers[i].ljust(widths[i]) for i in range(cols)
                ) + " |"

                output.write(f"\n[{stage}] Tabla {lo.Name} en hoja {ws.Name}:\n")
                output.write(sep_line + "\n")
                output.write(hdr_line + "\n")
                output.write(sep_line + "\n")
                if rows_list:
                    for row in rows_list:
                        output.write(
                            "| " + " | ".join(row[i].ljust(widths[i]) for i in range(cols)) + " |\n"
                        )
                else:
                    output.write(
                        "| " + " | ".join("".ljust(widths[i]) for i in range(cols)) + " |\n"
                    )
                output.write(sep_line + "\n")
        except Exception as e:
            output.write(f"[DEBUG] No se pudo procesar tablas en hoja {ws.Name}: {e}\n")


def rename_and_save_local(wb, original_name, output) -> Tuple[str, str]:
    """Guarda en LOCAL_OUTPUT_DIR con renombre por celda AA1."""
    raw = wb.Worksheets(1).Range("AA1").Value
    if hasattr(raw, 'strftime'):
        aa1_date_str = raw.strftime("%d-%m")
    else:
        try:
            aa1_date_str = datetime.strptime(str(raw), "%Y-%m-%d").strftime("%d-%m")
        except Exception:
            aa1_date_str = datetime.now().strftime("%d-%m")

    base, ext = os.path.splitext(original_name)
    if DATE_PATTERN.search(base):
        nuevo_base = DATE_PATTERN.sub(aa1_date_str, base)
    else:
        nuevo_base = f"{base} {aa1_date_str}"

    LOCAL_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    local_path = str(LOCAL_OUTPUT_DIR / f"{nuevo_base}{ext}")
    wb.SaveCopyAs(local_path)

    if os.path.exists(local_path) and os.path.getsize(local_path) > 0:
        output.write(f"[OK] Guardado LOCAL: {local_path}\n")
    else:
        output.write(f"[ERROR] Fallo al guardar LOCAL: {local_path}\n")

    return local_path, aa1_date_str


def extract_tablas_data(wb) -> Tuple[List[str], List[dict]]:
    """Extrae datos de tablas RESUMEN_KPI como dicts estructurados.
    Retorna (headers, rows) donde cada row es un dict clave->valor."""
    all_rows: List[dict] = []
    headers: List[str] = []
    for ws in wb.Worksheets:
        try:
            for lo in ws.ListObjects:
                if not lo.Name.upper().startswith("RESUMEN_KPI"):
                    continue
                hdr_rng = lo.HeaderRowRange
                cols = hdr_rng.Columns.Count
                headers = [str(hdr_rng.Cells(1, j).Value or "") for j in range(1, cols + 1)]
                data_rng = lo.DataBodyRange
                if data_rng:
                    for i in range(1, data_rng.Rows.Count + 1):
                        row = {}
                        for j in range(1, cols + 1):
                            cell = ws.Cells(
                                data_rng.Row + i - 1,
                                data_rng.Column + j - 1
                            ).Value
                            if hasattr(cell, "strftime"):
                                row[headers[j - 1]] = cell.strftime("%Y-%m-%d")
                            else:
                                row[headers[j - 1]] = cell
                        all_rows.append(row)
        except Exception:
            pass
    return headers, all_rows


def _worker_process(full_path: str, q: Queue):
    """
    Proceso hijo: procesa 1 Excel y envía resultado por Queue.

    Protocolo:
      1. Envía {"type": "excel_pid", "pid": ..., "stem": ...}
      2. Al finalizar: {"stem": ..., "resumen": {...}, "stdout": str}
    """
    t0 = time.time()
    output = StringIO()
    nombre = os.path.basename(full_path)
    stem   = Path(nombre).stem
    saved_local = ""
    aa1_date    = ""
    excel = None
    wb = None

    def wlog(msg: str):
        output.write(msg + "\n")

    try:
        wlog(f"{'='*60}")
        wlog(f"  Procesando: {nombre}")
        wlog(f"{'='*60}")

        # ── Abrir Excel ──
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 1  # msoAutomationSecurityLow

        # Reportar PID al padre
        excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
        q.put({"type": "excel_pid", "pid": excel_pid, "stem": stem})
        wlog(f"[PID] Excel instance: {excel_pid}")

        # ── Abrir workbook ──
        wb = excel.Workbooks.Open(
            full_path, False, False, None,
            EXCEL_PASSWORD, EXCEL_PASSWORD
        )
        wlog("[OK] Workbook abierto.")

        # ── Snapshot antes (texto + datos) ──
        print_resumen_kpi(wb, 'Before', output)
        tab_headers, tablas_before = extract_tablas_data(wb)

        # ── Refresh ──
        wlog("[..] Ejecutando RefreshAll…")
        disable_background_refresh(wb)
        wb.UpdateLinks = 1
        wb.RefreshAll()
        wait_until_queries_done(excel, wb, timeout_s=1800, poll_s=0.5)
        wlog("[OK] PowerQuery finalizado.")

        # ── PivotTables ──
        wlog("[..] Refrescando PivotTables…")
        pivot_count = 0
        for ws in wb.Worksheets:
            try:
                for pt in ws.PivotTables():
                    pt.RefreshTable()
                    pivot_count += 1
            except Exception as e:
                wlog(f"[WARN] PivotTables en hoja {ws.Name}: {e}")
        wlog(f"[OK] {pivot_count} PivotTables actualizadas.")

        # ── Snapshot después (texto + datos) ──
        print_resumen_kpi(wb, 'After', output)
        _, tablas_after = extract_tablas_data(wb)

        # ── Guardar copia local ──
        saved_local, aa1_date = rename_and_save_local(wb, nombre, output)

        # ── Cerrar ──
        try:
            wb.Close(False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass

        dur = round(time.time() - t0, 2)
        resumen = {
            "archivo": nombre,
            "ok": True,
            "ruta_local": saved_local,
            "fecha_AA1": aa1_date,
            "duracion_s": dur,
            "error": ""
        }
        q.put({"stem": stem, "resumen": resumen, "stdout": output.getvalue(),
               "tablas": {"headers": tab_headers, "before": tablas_before, "after": tablas_after}})

    except Exception as e:
        # ── Limpieza en error ──
        for obj in (wb, excel):
            try:
                if obj is not None:
                    obj.Close(False) if hasattr(obj, 'Close') else obj.Quit()
            except Exception:
                pass

        dur = round(time.time() - t0, 2)
        salida_err = output.getvalue() + f"\n[{nombre}] ERROR: {e}\n"
        resumen = {
            "archivo": nombre,
            "ok": False,
            "ruta_local": saved_local,
            "fecha_AA1": aa1_date,
            "duracion_s": dur,
            "error": str(e)
        }
        q.put({"stem": stem, "resumen": resumen, "stdout": salida_err})


# ─────────────────────── TABLA COMPARATIVA ACUMULADA ───────────────────────

def print_tablas_comparativa(tablas_coleccion: List[dict]):
    """Muestra tabla consolidada: compara Before vs After de todos los archivos.
    7 columnas: KPI | SUBÁREA | SERVICIO | ANTES | DESPUÉS | Δ GAP | ALCANCE
    Ordenada por SERVICIO (col 3). ALCANCE en rojo si cambió entre pasadas."""

    if not tablas_coleccion:
        return

    section("TABLA COMPARATIVA ACUMULADA")

    # Armar filas planas [archivo, kpi, subarea, servicio, antes, despues, alcance, gap]
    planas: List[tuple] = []
    col_kpi = 0
    col_sub = 1
    col_serv = 2
    col_res = 3
    col_alc = 4

    for item in tablas_coleccion:
        archivo = item["archivo"]
        hdr = item["headers"]
        bef = {r.get(hdr[col_serv], ""): r for r in item["before"]}
        aft_list = item["after"]

        for row_a in aft_list:
            key = str(row_a.get(hdr[col_serv], ""))
            row_b = bef.get(key, {})

            # Valores
            kpi = str(row_a.get(hdr[col_kpi], ""))
            sub = str(row_a.get(hdr[col_sub], ""))
            serv = key

            # RESULTADO (col 4) — numérico
            raw_antes = row_b.get(hdr[col_res])
            raw_despues = row_a.get(hdr[col_res])
            try:
                antes_f = float(raw_antes) if raw_antes is not None else None
            except (ValueError, TypeError):
                antes_f = None
            try:
                despues_f = float(raw_despues) if raw_despues is not None else None
            except (ValueError, TypeError):
                despues_f = None

            if antes_f is not None and despues_f is not None:
                gap = despues_f - antes_f
                gap_str = f"{gap:+.4f}"
                antes_str = f"{antes_f:.4f}"
                despues_str = f"{despues_f:.4f}"
            else:
                gap_str = "N/A"
                antes_str = str(raw_antes or "")
                despues_str = str(raw_despues or "")

            # ALCANCE (col 5) — PASA/NO PASA
            alc_before = str(row_b.get(hdr[col_alc], "")).strip().lower()
            alc_after = str(row_a.get(hdr[col_alc], "")).strip()
            alc_changed = (alc_before != alc_after.lower() and alc_before != "")

            planas.append((archivo, kpi, sub, serv, antes_str, despues_str,
                          gap_str, alc_after, alc_changed))

    if not planas:
        info("(sin datos de RESUMEN_KPI para comparar)")
        return

    # Ordenar por SERVICIO (col 3, índice 3 en la tupla)
    planas.sort(key=lambda r: (r[3], r[0], r[1]))

    # Calcular anchos
    ancho_archivo = max(len(p[0]) for p in planas)
    ancho_kpi = max(len(p[1]) for p in planas)
    ancho_sub = max(len(p[2]) for p in planas)
    ancho_serv = max(len(p[3]) for p in planas)
    ancho_antes = max(len(p[4]) for p in planas)
    ancho_despues = max(len(p[5]) for p in planas)
    ancho_gap = max(len(p[6]) for p in planas)
    ancho_alc = max(len(p[7]) for p in planas)

    cols_h = ["ARCHIVO", "KPI", "SUBÁREA", "SERVICIO", "ANTES", "DESPUÉS", "Δ GAP", "ALCANCE"]
    col_widths = [
        max(ancho_archivo, len(cols_h[0])),
        max(ancho_kpi, len(cols_h[1])),
        max(ancho_sub, len(cols_h[2])),
        max(ancho_serv, len(cols_h[3])),
        max(ancho_antes, len(cols_h[4])),
        max(ancho_despues, len(cols_h[5])),
        max(ancho_gap, len(cols_h[6])),
        max(ancho_alc, len(cols_h[7])),
    ]

    def fmt_row(vals, color_override=None):
        parts = []
        for i, v in enumerate(vals):
            padded = str(v).ljust(col_widths[i])
            if color_override:
                parts.append(_c(color_override, padded))
            else:
                parts.append(padded)
        return " │ ".join(parts)

    sep_line = "─" * (sum(col_widths) + 3 * (len(col_widths) - 1) + 2)

    print(f"  {_c(_C_GREEN, '┌' + '─' * (sep_line.count('─') + 2) + '┐')}")
    print(f"  {_c(_C_GREEN, '│')}  {_c(_C_CYAN, 'Tabla consolidada — todas las tablas RESUMEN_KPI ordenadas por SERVICIO')}")
    print(f"  {_c(_C_GREEN, '├' + '─' * (sep_line.count('─') + 2) + '┤')}")
    print(f"  {_c(_C_GREEN, '└' + '─' * (sep_line.count('─') + 2) + '┘')}")
    print()

    # Header
    print(f"  {_c(_C_CYAN, fmt_row(cols_h))}")
    print(f"  {_c(_C_GRAY, sep_line)}")

    last_file = None
    for idx, p in enumerate(planas):
        arch, kpi, sub, serv, antes, despues, gap_str, alc_text, alc_changed = p

        # Separador entre archivos
        if arch != last_file and last_file is not None:
            print(f"  {_c(_C_GRAY, sep_line)}")
        last_file = arch

        # Armar fila
        vals = [arch, kpi, sub, serv, antes, despues, gap_str, alc_text]

        marcar_rojo = False
        color_overrides = [None] * 8

        # Colorear gap: verde si positivo, rojo si negativo
        if gap_str not in ("N/A", "0.0000", ""):
            try:
                g = float(gap_str)
                if g > 0:
                    color_overrides[6] = _C_GREEN
                elif g < 0:
                    color_overrides[6] = _C_RED
            except ValueError:
                pass

        # ALCANCE en rojo si cambió
        if alc_changed:
            color_overrides[7] = _C_RED
            marcar_rojo = True

        # Marcador visual si hubo cambios
        marker = _c(_C_YELLOW, "⚡") if marcar_rojo or gap_str not in ("N/A", "0.0000", "") else " "

        line_parts = []
        for i, v in enumerate(vals):
            padded = str(v).ljust(col_widths[i])
            c = color_overrides[i]
            line_parts.append(_c(c, padded) if c else padded)

        line = " │ ".join(line_parts)
        print(f"  {marker} {line}")

    print(f"  {_c(_C_GRAY, sep_line)}")

    # Leyenda
    print()
    info(f"{_c(_C_GREEN, 'verde')} = gap positivo (mejoró)")
    info(f"{_c(_C_RED, 'rojo')}   = gap negativo o ALCANCE cambió")
    info(f"{_c(_C_YELLOW, '⚡')}  = fila con cambios")
    print()
    info(f"Total filas: {len(planas)} de {len(tablas_coleccion)} archivos")
    print()


# ────────────────────────────── ORQUESTACIÓN PADRE ──────────────────────────────
def main():
    section("INICIO — ExelenParalelo")
    info(f"Origen:   {INPUT_EXCEL_DIR}")
    info(f"Destino:  {LOCAL_OUTPUT_DIR}")
    info(f"Workers:  {MAX_WORKERS} | Timeout: {TIMEOUT_S}s | Reintentos: {MAX_RETRIES}")
    info(f"CSV log:  {CSV_PROCESS_PATH}")
    print()

    # ── Cerrar Excels residuales ──
    if CLOSE_ALL_EXCEL_BEFORE:
        info("Cerrando instancias previas de Excel…")
        close_all_excels_before_start(force=True)
        time.sleep(CLOSE_WAIT_S)

    LOCAL_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    append_csv_header_if_missing(CSV_PROCESS_PATH, CSV_PROCESS_HEADER)

    # ── Escanear archivos ──
    files = glob.glob(os.path.join(INPUT_EXCEL_DIR, "*.xls*"))
    if not files:
        warn("No hay archivos para procesar.")
        return 0

    ok(f"Archivos detectados: {len(files)}")
    print()

    # ── Pool state ──
    pendientes: List[Tuple[str, int]] = [(f, MAX_RETRIES) for f in files]
    # stem -> (proc, queue, start_ts, full_path, intento_idx, excel_pid, pre_msg)
    activos: Dict[str, Tuple[Process, Queue, float, str, int, Optional[int], Optional[dict]]] = {}
    resultados: List[Tuple[str, bool]] = []
    detalles_finales: Dict[str, dict] = {}
    tablas_coleccion: List[dict] = []  # {archivo, headers, before, after}
    t0 = time.time()

    # ── Barra de progreso ──
    def _barra():
        done = len(resultados)
        active = len(activos)
        pending = len(pendientes)
        total = done + active + pending
        if total == 0:
            return
        elapsed = time.time() - t0
        pct = done / total
        ancho = 20
        fill = int(ancho * pct)
        bar = "█" * fill + "░" * (ancho - fill)
        sys.stdout.write(
            f"\r  {_c(_C_CYAN, '📊')} {bar} "
            f"{done}/{total}  "
            f"({active} activos, {pending} pendientes)  "
            f"{elapsed:6.0f}s  ({pct*100:5.1f}%)   "
        )
        sys.stdout.flush()

    # ── Lanzador de workers ──
    def lanzar():
        if not pendientes or len(activos) >= MAX_WORKERS:
            return
        full_path, retries_left = pendientes.pop(0)
        stem = Path(full_path).stem
        intento_idx = (MAX_RETRIES - retries_left) + 1

        q = Queue(maxsize=4)
        p = Process(target=_worker_process, args=(full_path, q), daemon=True)
        p.start()

        excel_pid = None
        pre_msg: Optional[dict] = None

        # Ventana de 10s para capturar PID o resultado express
        deadline = time.time() + 10
        while time.time() < deadline:
            try:
                m = q.get_nowait()
                if isinstance(m, dict) and m.get("type") == "excel_pid":
                    excel_pid = int(m["pid"])
                elif isinstance(m, dict) and "resumen" in m:
                    pre_msg = m
                    break
            except queue.Empty:
                time.sleep(0.05)
            except Exception:
                break

        activos[stem] = (p, q, time.time(), full_path, intento_idx, excel_pid, pre_msg)
        print()
        info(f"▶ {stem}  (PID {p.pid}, intento {intento_idx}/{MAX_RETRIES})")
        _barra()

    # ── Carga inicial ──
    for _ in range(min(MAX_WORKERS, len(pendientes))):
        lanzar()
    print()

    # ── Bucle principal de vigilancia ──
    while activos or pendientes:
        for stem in list(activos.keys()):
            p, q, start_ts, full_path, intento_idx, excel_pid, pre_msg = activos[stem]
            elapsed = time.time() - start_ts
            nombre = os.path.basename(full_path)

            # ── Worker terminó ──
            if not p.is_alive():
                # Drenar cola
                mensajes = []
                if pre_msg is not None:
                    mensajes.append(pre_msg)
                try:
                    while True:
                        mensajes.append(q.get_nowait())
                except queue.Empty:
                    pass
                except Exception:
                    pass

                # Capturar PID rezagado
                for m in mensajes:
                    if isinstance(m, dict) and m.get("type") == "excel_pid":
                        try:
                            excel_pid = int(m["pid"])
                        except Exception:
                            pass

                # Buscar resumen
                msg = next(
                    (m for m in mensajes if isinstance(m, dict) and "resumen" in m),
                    None
                )
                if msg is None:
                    msg = {
                        "stem": stem,
                        "resumen": {
                            "archivo": nombre, "ok": False,
                            "ruta_local": "", "fecha_AA1": "",
                            "duracion_s": round(elapsed, 2),
                            "error": "Proceso finalizado sin resumen"
                        },
                        "stdout": ""
                    }

                # Volcar stdout del worker (indentado con ┊)
                if msg.get("stdout"):
                    for line in msg["stdout"].splitlines():
                        if line.strip():
                            print(f"  {_c(_C_GRAY, '┊')} {line}")

                resumen = msg["resumen"]
                detalles_finales[stem] = resumen
                resultados.append((stem, resumen["ok"]))

                # Coleccionar datos de tablas RESUMEN_KPI
                if msg.get("tablas") and msg["tablas"].get("headers"):
                    tablas_coleccion.append({
                        "archivo": resumen["archivo"],
                        "headers": msg["tablas"]["headers"],
                        "before": msg["tablas"].get("before", []),
                        "after": msg["tablas"].get("after", []),
                    })

                if resumen["ok"]:
                    ok(f"{stem} → {resumen['ruta_local']} ({resumen['duracion_s']}s)")
                else:
                    error(f"{stem} → {resumen['error']} ({resumen['duracion_s']}s)")

                # CSV
                append_csv_row(
                    CSV_PROCESS_PATH,
                    [
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        str(intento_idx),
                        resumen["archivo"],
                        str(resumen["ok"]),
                        resumen["ruta_local"],
                        resumen["fecha_AA1"],
                        str(resumen["duracion_s"]),
                        (resumen["error"] or "").replace("\n", " ").replace(";", ","),
                    ],
                )

                # Cleanup y siguiente
                try:
                    p.join(timeout=1)
                except Exception:
                    pass
                del activos[stem]
                lanzar()
                _barra()

            # ── Timeout ──
            elif elapsed > TIMEOUT_S:
                # Drenar mensajes tardíos
                try:
                    while True:
                        m = q.get_nowait()
                        if isinstance(m, dict) and m.get("type") == "excel_pid":
                            try:
                                excel_pid = int(m["pid"])
                                activos[stem] = (
                                    p, q, start_ts, full_path,
                                    intento_idx, excel_pid, pre_msg
                                )
                            except Exception:
                                pass
                        elif isinstance(m, dict) and "resumen" in m:
                            pre_msg = m
                            activos[stem] = (
                                p, q, start_ts, full_path,
                                intento_idx, excel_pid, pre_msg
                            )
                except queue.Empty:
                    pass
                except Exception:
                    pass

                warn(f"TIMEOUT — {stem} superó {TIMEOUT_S}s")
                try:
                    p.terminate()
                    p.join(timeout=5)
                except Exception:
                    pass

                if excel_pid:
                    try:
                        subprocess.run(
                            ["taskkill", "/PID", str(excel_pid), "/F", "/T"],
                            capture_output=True, text=True, timeout=10
                        )
                        warn(f"Excel PID {excel_pid} eliminado.")
                    except Exception as e:
                        warn(f"No se pudo matar Excel {excel_pid}: {e}")

                append_csv_row(
                    CSV_PROCESS_PATH,
                    [
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        str(intento_idx), nombre, "False",
                        "", "",
                        str(round(elapsed, 2)),
                        f"Timeout {TIMEOUT_S}s",
                    ],
                )

                retries_left = MAX_RETRIES - intento_idx + 1
                if retries_left > 0:
                    info(f"↻ Reintentando {stem} ({retries_left - 1} restantes)")
                    pendientes.insert(0, (full_path, retries_left - 1))
                else:
                    resultados.append((stem, False))
                    detalles_finales[stem] = {
                        "archivo": nombre, "ok": False,
                        "ruta_local": "", "fecha_AA1": "",
                        "duracion_s": round(elapsed, 2),
                        "error": f"Timeout {TIMEOUT_S}s"
                    }

                del activos[stem]
                lanzar()
                _barra()

        time.sleep(0.5)

    # ──────────────────────────────── RESUMEN FINAL ────────────────────────────────
    print("\n")
    section("RESUMEN FINAL")

    exitosos = [r for r in resultados if r[1]]
    errores  = [r for r in resultados if not r[1]]
    total_dur = round(time.time() - t0, 2)

    ok(f"Procesados: {len(resultados)}")
    ok(f"Éxitos:     {len(exitosos)}")
    if errores:
        error(f"Errores:    {len(errores)}")
        for stem, _ in errores:
            det = detalles_finales.get(stem, {})
            print(f"            • {stem}: {det.get('error', '?')}")
    else:
        ok("Errores:     0")

    print()
    info(f"Duración total: {total_dur:.2f}s")
    info(f"CSV:           {CSV_PROCESS_PATH}")
    print()

    # Tabla comparativa acumulada
    print_tablas_comparativa(tablas_coleccion)

    return 0 if len(errores) == 0 else 1


if __name__ == "__main__":
    rc = 1
    try:
        rc = main()
    finally:
        if CLOSE_ALL_EXCEL_AFTER:
            info("Cerrando instancias de Excel residuales…")
            close_all_excels_before_start(force=True)
            time.sleep(CLOSE_WAIT_S)
    sys.exit(rc)
