#!/usr/bin/env python3
import argparse, re, sys
from pathlib import Path
import pandas as pd
import numpy as np

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

def is_date_like(s):
    if s is None or (isinstance(s,float) and np.isnan(s)):
        return False
    s = str(s).strip()
    return bool(DATE_RE.match(s))

def parse_date(s):
    try:
        return pd.to_datetime(str(s).strip(), dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT

def parse_number(raw):
    """Devuelve (value_num, is_percent). No rompe si no es número."""
    if raw is None:
        return None, False
    s = str(raw).strip()
    is_pct = s.endswith("%")
    s_clean = s[:-1] if is_pct else s
    # quita miles '.', cambia coma decimal a punto
    s_clean = s_clean.replace(".", "").replace(",", ".")
    # vacío o no numérico => None
    try:
        val = float(s_clean)
        if is_pct:
            val = val/100.0
        return val, is_pct
    except Exception:
        return None, is_pct

def find_date_header_row(df):
    """Busca la fila con más fechas tipo 1-1-25, devuelve (row_idx, first_col, last_col, total_col)."""
    best = (-1, -1, -1, -1)  # (row, first_date_col, last_date_col, total_col)
    for r in range(len(df)):
        row = df.iloc[r]
        date_cols = [c for c, v in enumerate(row) if is_date_like(v)]
        if len(date_cols) >= 5:
            # intenta hallar "TOTAL" a la derecha
            tot_col = -1
            for c in range(max(date_cols)+1, len(row)):
                v = row.iloc[c]
                if isinstance(v, str) and v.strip().upper() == "TOTAL":
                    tot_col = c
                    break
            first = min(date_cols); last = max(date_cols)
            best = (r, first, last, tot_col)
            # nos quedamos con la primera “buena” que tenga al menos 5 fechas
            break
    return best


    # -------------------- Helpers-END ----------------------------
def normalize_file(path: Path, sheet_name: str = "Daily"):
    # Lee Excel/CSV preservando texto crudo
    if path.suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
        df = pd.read_excel(
            path,
            sheet_name=sheet_name,  # <<--- usa la hoja "Daily"
            header=None,
            dtype=object,
            engine="openpyxl",
        )
    elif path.suffix.lower() == ".csv":
        df = pd.read_csv(path, header=None, dtype=object)
    else:
        raise ValueError(f"Formato no soportado: {path.suffix}")

    date_hdr_row, first_c, last_c, total_c = find_date_header_row(df)
    if date_hdr_row < 0:
        raise RuntimeError(f"No encontré fila de fechas en {path.name}")

    # ... (resto de la lógica igual)
    # mapeo de fechas, barrido de filas, records, etc.
    # retorna DataFrame 'out'

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, help="Ruta del CSV maestro (formato largo)")
    ap.add_argument("--sheet", default="Daily", help="Nombre de la hoja dentro del Excel")
    ap.add_argument("files", nargs="+", help="Rutas a .xlsx/.xls/.csv")
    args = ap.parse_args()

    frames = []
    for f in args.files:
        df = normalize_file(Path(f), sheet_name=args.sheet)
        frames.append(df)

    master = pd.concat(frames, ignore_index=True)
    cols = ["source_file","section","metric_label","date","value_raw","value_num","is_percent","is_total"]
    master = master.reindex(columns=cols)
    master.to_csv(args.out, index=False)
    by_file = master.groupby("source_file")["value_raw"].count().to_dict()
    print("OK. Filas por archivo:", by_file)
    print("Salida:", args.out)

if __name__ == "__main__":
    sys.exit(main())

