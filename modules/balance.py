# -*- coding: utf-8 -*-
from __future__ import annotations
from dataclasses import dataclass
from typing import List, Tuple, Optional
import re
import pandas as pd
from .writer import DiarioWriter, normalize_value

# ==========================
# Config y sinónimos
# ==========================
HEADER_SCAN_ROWS = 50  # cuántas filas iniciales escanear para hallar encabezado

ACCOUNT_SYNONYMS = {
    "cuenta", "cuentas", "concepto", "conceptos", "plan de cuentas", "nombre de cuenta"
}

OPENING_SYNONYMS = {
    "apertura", "saldo inicial", "saldo al inicio", "saldo inicio",
    "inicial", "inicio", "saldo de apertura"
}

MONTH_MAP = {
    "enero":1, "ene":1, "jan":1,
    "febrero":2, "feb":2,
    "marzo":3, "mar":3,
    "abril":4, "abr":4,
    "mayo":5, "may":5,
    "junio":6, "jun":6,
    "julio":7, "jul":7,
    "agosto":8, "ago":8,
    "septiembre":9, "setiembre":9, "sep":9, "set":9,
    "octubre":10, "oct":10,
    "noviembre":11, "nov":11,
    "diciembre":12, "dic":12, "dec":12,
}

def _norm(s) -> str:
    return str(s).strip().lower()

# ==========================
# Detección de meses
# ==========================
def _month_num_from_header(h: str) -> Optional[int]:
    hs = _norm(h)

    # 1) Nombre del mes (ene, enero, dic/19, etc.)
    for key, num in MONTH_MAP.items():
        if re.search(rf"\b{re.escape(key)}\b", hs):
            return num

    # 2) yyyy-mm / yyyy/mm / yyyy mm
    m = re.search(r"\b(20\d{2})[-_/ ](0?[1-9]|1[0-2])\b", hs)
    if m:
        return int(m.group(2))

    # 3) mm-yyyy / mm/yy / mm yy
    m = re.search(r"\b(0?[1-9]|1[0-2])[-_/ ](20\d{2}|\d{2})\b", hs)
    if m:
        return int(m.group(1))

    # 4) Solo número 1..12, evitando palabras contables
    short = re.fullmatch(r"(0?[1-9]|1[0-2])", hs.replace(" ", ""))
    if short and not any(t in hs for t in ("debe", "haber", "saldo", "apertura", "inicial", "inicio", "total")):
        return int(short.group(1))

    return None

def _looks_like_opening(h: str) -> bool:
    hs = _norm(h)
    return any(s in hs for s in OPENING_SYNONYMS)

# ==========================
# Detección de fila de encabezados
# ==========================
def _row_header_score(row_vals: List[str]) -> float:
    """
    Puntúa una fila: cuenta + cantidad de meses + penaliza números puros.
    Se usa para elegir la fila que "mejor" parece encabezado.
    """
    txts = [str(x) for x in row_vals]
    if not any(t.strip() for t in txts):
        return -1.0

    tokens = [_norm(t) for t in txts if str(t).strip() != ""]
    if not tokens:
        return -1.0

    has_account = any(any(k in tok for k in ACCOUNT_SYNONYMS) for tok in tokens)
    month_hits = sum(1 for t in tokens if _month_num_from_header(t) is not None)

    # penalizar si casi todos son números/NaN
    numeric_like = 0
    for t in tokens:
        try:
            float(t.replace(".", "").replace(",", "."))
            numeric_like += 1
        except:
            pass

    score = 0.0
    if has_account:
        score += 3.0
    score += 1.2 * month_hits
    score -= 0.5 * numeric_like
    return score

def _load_balance_table(input_path: str, sheet_name: str, header_row: Optional[int]) -> pd.DataFrame:
    """
    Si header_row (0-based) viene dado por la app, lo usamos.
    Si no, detectamos automáticamente escaneando las primeras HEADER_SCAN_ROWS filas.
    """
    if header_row is not None:
        df = pd.read_excel(input_path, sheet_name=sheet_name, header=header_row, dtype=object)
        return df.dropna(axis=1, how="all").dropna(axis=0, how="all")

    df_raw = pd.read_excel(input_path, sheet_name=sheet_name, header=None, dtype=object)
    max_scan = min(HEADER_SCAN_ROWS, len(df_raw))
    best_row = 0
    best_score = -1e9
    for r in range(max_scan):
        row_vals = [df_raw.iloc[r, c] for c in range(df_raw.shape[1])]
        sc = _row_header_score(row_vals)
        if sc > best_score:
            best_score = sc
            best_row = r

    # Reasignar encabezados a partir de best_row
    new_cols = list(df_raw.iloc[best_row].values)
    df = df_raw.iloc[best_row + 1:].copy()
    df.columns = new_cols
    return df.dropna(axis=1, how="all").dropna(axis=0, how="all")

# ==========================
# Especificación del balance
# ==========================
@dataclass
class BalanceSpec:
    account_col: str
    opening_cols: List[str]
    month_cols: List[Tuple[int, str]]

def _pick_account_col(df: pd.DataFrame) -> str:
    cols = list(df.columns)

    # 1) por sinónimos en el nombre
    for c in cols:
        if any(s in _norm(c) for s in ACCOUNT_SYNONYMS):
            return c

    # 2) heurística: columna con más "texto" (menos numérica)
    best_col = cols[0]
    best_score = -1e9
    for c in cols:
        s = df[c]
        non_null = s.notna().sum()
        numeric_like = 0
        for v in s.dropna().head(80):
            try:
                float(str(v).replace(".", "").replace(",", "."))
                numeric_like += 1
            except:
                pass
        score = non_null - 2 * numeric_like
        if score > best_score:
            best_score = score
            best_col = c
    return best_col

def detect_balance_spec(df: pd.DataFrame) -> BalanceSpec:
    cols = list(df.columns)

    account_col = _pick_account_col(df)

    opening_cols = [c for c in cols if _looks_like_opening(c)]

    month_cols: List[Tuple[int, str]] = []
    for c in cols:
        n = _month_num_from_header(c)
        if n is not None:
            month_cols.append((n, str(c)))

    # ordenar y quitar duplicados por mes
    seen = set()
    ordered = []
    for m, name in sorted(month_cols, key=lambda x: x[0]):
        if m not in seen:
            ordered.append((m, name))
            seen.add(m)

    return BalanceSpec(account_col, opening_cols, ordered)

# ==========================
# Construcción de líneas
# ==========================
def _find_debe_haber_pair(df: pd.DataFrame, base_header: str) -> Optional[Tuple[str, str]]:
    """
    Busca columnas emparejadas Debe/Haber alrededor de 'base_header':
      <base> Debe / <base> Haber
      <base>-Debe / <base>-Haber
      Debe <base> / Haber <base>
      variantes con '/', '_' o pegadas
    """
    cols = [str(c) for c in df.columns]
    b = str(base_header).strip()

    pat_debe = [
        rf"^{re.escape(b)}\s*[-_/]?\s*debe$",
        rf"^debe\s*[-_/]?\s*{re.escape(b)}$",
        rf"^{re.escape(b)}debe$",
        rf"^debe{re.escape(b)}$",
    ]
    pat_haber = [
        rf"^{re.escape(b)}\s*[-_/]?\s*haber$",
        rf"^haber\s*[-_/]?\s*{re.escape(b)}$",
        rf"^{re.escape(b)}haber$",
        rf"^haber{re.escape(b)}$",
    ]

    debe = None; haber = None
    for c in cols:
        cl = _norm(c)
        if any(re.fullmatch(p, cl) for p in pat_debe):
            debe = c if debe is None else debe
        if any(re.fullmatch(p, cl) for p in pat_haber):
            haber = c if haber is None else haber

    if debe and haber:
        return (debe, haber)
    return None

def build_opening_lines(df: pd.DataFrame, spec: BalanceSpec):
    if not spec.opening_cols:
        return []
    lines = []
    for _, row in df.iterrows():
        cuenta = str(row.get(spec.account_col, "")).strip()
        if not cuenta:
            continue
        total = 0.0
        for c in spec.opening_cols:
            total += normalize_value(row.get(c, 0))
        if abs(total) < 1e-12:
            continue
        debe = total if total > 0 else 0.0
        haber = -total if total < 0 else 0.0
        lines.append((cuenta, debe, haber))
    return lines

def build_month_lines(df: pd.DataFrame, spec: BalanceSpec, month_col: str):
    """
    Soporta:
      A) una sola columna por mes (valor +/- => Debe/Haber)
      B) par <Mes> Debe / <Mes> Haber (con -, /, _, o pegados, o invertidos)
    """
    pair = _find_debe_haber_pair(df, month_col)
    lines: List[Tuple[str, float, float]] = []

    if pair:
        cde, cha = pair
        for _, row in df.iterrows():
            cuenta = str(row.get(spec.account_col, "")).strip()
            if not cuenta:
                continue
            d = normalize_value(row.get(cde, 0))
            h = normalize_value(row.get(cha, 0))
            if abs(d) < 1e-12 and abs(h) < 1e-12:
                continue
            lines.append((cuenta, d, h))
        return lines

    # Columna única con signo
    for _, row in df.iterrows():
        cuenta = str(row.get(spec.account_col, "")).strip()
        if not cuenta:
            continue
        val = normalize_value(row.get(month_col, 0))
        if abs(val) < 1e-12:
            continue
        debe = val if val > 0 else 0.0
        haber = -val if val < 0 else 0.0
        lines.append((cuenta, debe, haber))
    return lines

def month_title(m: int, year: int) -> str:
    long_names = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
    label = long_names.get(m, f"Mes {m}")
    return f"As. Movimientos {label} {year}"

# ==========================
# Entrada principal
# ==========================
def build_diario_from_balance(input_path: str, balance_sheet: str|None, year: int, header_row: Optional[int] = None):
    xls = pd.ExcelFile(input_path)
    bsheet = balance_sheet if (balance_sheet and balance_sheet in xls.sheet_names) else xls.sheet_names[0]

    # Cargar tabla normalizada (con encabezados reales)
    df_balance = _load_balance_table(input_path, bsheet, header_row)

    spec = detect_balance_spec(df_balance)

    writer = DiarioWriter(year)
    last_n = 0

    # Apertura
    apertura = build_opening_lines(df_balance, spec)
    if apertura:
        last_n += 1
        writer.header(last_n)
        total_d = total_h = 0.0
        for cta, d, h in apertura:
            writer.append_row({writer.COL_B: cta, writer.COL_D: d, writer.COL_E: h})
            total_d += d; total_h += h
        writer.totales(f"Asiento de Apertura {year}", total_d, total_h)
        writer.blank_row()

    # Mensuales en orden
    for m, col in spec.month_cols:
        lines = build_month_lines(df_balance, spec, col)
        if not lines:
            continue
        last_n += 1
        writer.header(last_n)
        total_d = total_h = 0.0
        for cta, d, h in lines:
            writer.append_row({writer.COL_B: cta, writer.COL_D: d, writer.COL_E: h})
            total_d += d; total_h += h
        writer.totales(month_title(m, year), total_d, total_h)
        writer.blank_row()

    return writer, last_n
