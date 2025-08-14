#!/usr/bin/env python3
import argparse, re, sys
from pathlib import Path
import pandas as pd
import numpy as np
import datetime as _dt
import math


# -------------------- Helpers ----------------------------
DATE_RE = re.compile(r"^\s*(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\s*$", re.I)

SECTION_SYNONYMS = {
    "RESTAURANTE":"RESTAURANT", "RESTAURANT":"RESTAURANT",
    "BAR":"BAR", "BER":"BAR",
    "BODA":"BANQUETING","BENQUETING":"BANQUETING","BANQUETING":"BANQUETING",
    "EMPRESA":"MICE","MICE":"MICE",
    "PARTICULAR":"INDIVIDUALS","INDIVIDUALS":"INDIVIDUALS",
    "TIENDA RESTAURANTE 43":"SHOP","SHOP":"SHOP",
    "WALK IN":"WALKIN","WALKIN":"WALKIN",
    "INTERNO":"EMPLOYEES","EMPLEADOS":"EMPLOYEES","Empleados":"EMPLOYEES",
}

def is_date_like(x):
    """True si x parece una fecha: Timestamp, date/datetime o string parseable (DD/MM/YY admite dayfirst)."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return False
    if isinstance(x, (pd.Timestamp, _dt.date, _dt.datetime)):
        return True
    s = str(x).strip()
    if not s:
        return False
    d = pd.to_datetime(s, dayfirst=True, errors="coerce")
    return pd.notna(d) and (1900 <= d.year <= 2100)  # evita falsos positivos

ISO_RE = re.compile(r"^\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2}:\d{2})?$")

def parse_date(s):
    s = str(s).strip()
    if not s:
        return pd.NaT
    # Si es ISO (YYYY-MM-DD [HH:MM:SS]) no uses dayfirst
    if ISO_RE.match(s):
        return pd.to_datetime(s, errors="coerce")
    # Resto: formato europeo
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def parse_number(x):
    """
    Devuelve (valor: float|None, is_percent: bool).
    - No elimina puntos/decimales.
    - Acepta notación científica (p.ej. 1.47e-4).
    - Si el string acaba en '%', divide entre 100 y marca is_percent=True.
    - Si no es numérico, devuelve (None, False).
    """
    if x is None:
        return None, False
    if isinstance(x, float) and math.isnan(x):
        return None, False
    if isinstance(x, (int, float)):
        # Ya viene numérico de pandas/openpyxl
        return float(x), False

    s = str(x).strip()
    if not s or s.lower() in {"nan", "none", "-"}:
        return None, False

    is_pct = s.endswith("%")
    if is_pct:
        s = s[:-1].strip()

    # Normaliza espacios y decimal coma → punto. ¡No quites puntos!
    s = s.replace("\u00A0", "").replace(" ", "").replace(",", ".")

    try:
        val = float(s)  # Python acepta 1e-6, 3.4E+5, etc.
        # (notación científica soportada por el lenguaje). 
        # https://docs.python.org/3/reference/lexical_analysis.html#floating-point-literals
    except ValueError:
        return None, is_pct

    if is_pct:
        val /= 100.0  # Excel muestra 139%, el valor base es 1.39 → guardamos 0.0139 si el texto traía '%'
    return val, is_pct

def find_date_header_row(df):
    """
    Busca la fila con más 'celdas-fecha'. Admite fechas reales de Excel y strings '1-1-25'/'1/1/25'.
    Devuelve (row_idx, first_col, last_col, total_col).
    """
    best = (-1, -1, -1, -1)
    for r in range(len(df)):
        row = df.iloc[r]
        date_cols = [c for c, v in enumerate(row) if is_date_like(v)]
        if len(date_cols) >= 3:  # umbral prudente
            tot_col = -1
            for c in range(max(date_cols) + 1, len(row)):
                v = row.iloc[c]
                if isinstance(v, str) and v.strip().upper() in {"TOTAL", "TOTAL "}:
                    tot_col = c
                    break
            first = min(date_cols); last = max(date_cols)
            best = (r, first, last, tot_col)
            break
    return best
# -------------------- Helpers-END ----------------------------

def normalize_file(path: Path, sheet_name: str = "Daily") -> pd.DataFrame:
    # Lee Excel/CSV preservando texto crudo
    if path.suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
        df = pd.read_excel(
            path,
            sheet_name=sheet_name,  # hoja "Daily"
            header=None,
            dtype=object,          # preserva tipos crudos
            engine="openpyxl",
        )
    elif path.suffix.lower() == ".csv":
        df = pd.read_csv(path, header=None, dtype=object)
    else:
        raise ValueError(f"Formato no soportado: {path.suffix}")

    date_hdr_row, first_c, last_c, total_c = find_date_header_row(df)
    if date_hdr_row < 0:
        raise RuntimeError(f"No encontré fila de fechas en {path.name}")

    # Mapa col -> fecha (solo columnas con fecha válida)
    col2date = {}
    for c in range(first_c, last_c + 1):
        d = parse_date(df.iat[date_hdr_row, c])
        if pd.notna(d):
            col2date[c] = d.normalize()  # 00:00

    records = []
    current_section = None

    # Recorre filas por debajo del header de fechas
    for r in range(date_hdr_row + 1, len(df)):
        row = df.iloc[r]

        # Busca la etiqueta a la izquierda de la primera fecha (última no vacía)
        label = None
        for c in range(0, first_c):
            v = row.iloc[c]
            if isinstance(v, str) and v.strip() != "":
                label = str(v).strip()

        # Actualiza sección si corresponde (y normaliza nombre)
        if label:
            up = label.upper()
            if up in SECTION_SYNONYMS:
                current_section = SECTION_SYNONYMS[up]
                # Esta fila suele ser título de sección; puede que no tenga datos
                # Continuamos a mapear valores igualmente por si acaso.

        metric_label = None
        # Si label NO es una sección conocida, trátalo como métrica (COMIDA, BEBIDA, TRUE/FALSE, etc.)
        if label and (label.upper() not in SECTION_SYNONYMS):
            metric_label = label

        # Emite celdas diarias
        for c, d in col2date.items():
            v = row.iloc[c]
            if v is None or (isinstance(v, float) and np.isnan(v)):
                continue
            s = str(v).strip()
            if s == "":
                continue
            v = row.iloc[c]
            if is_date_like(v):
                continue  # <- evita filas-basura con "2025-01-xx 00:00:00"

            val_num, is_pct = parse_number(s)
            records.append({
                "source_file": path.name,
                "section": current_section,
                "metric_label": metric_label,
                "date": d,
                "value_raw": s,
                "value_num": val_num,
                "is_percent": is_pct,
                "is_total": False,
            })

        # Emite TOTAL si existe
        if total_c != -1 and total_c < len(row):
            tv = row.iloc[total_c]
            if tv is not None and not (isinstance(tv, float) and np.isnan(tv)) and str(tv).strip() != "":
                s = str(tv).strip()
                val_num, is_pct = parse_number(s)
                records.append({
                    "source_file": path.name,
                    "section": current_section,
                    "metric_label": metric_label,
                    "date": pd.NaT,       # totales sin fecha
                    "value_raw": s,
                    "value_num": val_num,
                    "is_percent": is_pct,
                    "is_total": True,
                })

    out = pd.DataFrame.from_records(records)
    # Ordena columnas y devuelve
    cols = ["source_file","section","metric_label","date","value_raw","value_num","is_percent","is_total"]
    out = out.reindex(columns=cols)
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, help="Ruta del CSV maestro (formato largo)")
    ap.add_argument("--sheet", default="Daily", help="Nombre de la hoja dentro del Excel")
    ap.add_argument("--peek", action="store_true", help="Guarda _peek_*.csv con las primeras 30 filas")
    ap.add_argument("files", nargs="+", help="Rutas a .xlsx/.xls/.csv")
    args = ap.parse_args()

    frames = []
    for f in args.files:
        df = normalize_file(Path(f), sheet_name=args.sheet)
        if args.peek:
            # guarda vista previa por archivo
            prev = df.head(30)
            prev.to_csv(f"_peek_{Path(f).stem}.csv", index=False)
        frames.append(df)

    master = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    cols = ["source_file","section","metric_label","date","value_raw","value_num","is_percent","is_total"]
    master = master.reindex(columns=cols)
    master.to_csv(args.out, index=False)
    by_file = master.groupby("source_file")["value_raw"].count().to_dict() if not master.empty else {}
    print("OK. Filas por archivo:", by_file)
    print("Salida:", args.out)

if __name__ == "__main__":
    sys.exit(main())
