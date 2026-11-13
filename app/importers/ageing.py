"""Ageing workbook -> Full Ageing + Account Analyst List."""
from python_calamine import load_workbook

from .common import CONVERT
from .specs import TABLES


def _header_row(rows, first_label):
    for i, r in enumerate(rows[:15]):
        if r and str(r[0]).strip() == first_label:
            return i
    raise ValueError(f"Header row starting with '{first_label}' not found.")


def _table(rows, table_id, first_label, width):
    h = _header_row(rows, first_label)
    cols = TABLES[table_id]["cols"]
    out = []
    for r in rows[h + 1:]:
        r = list(r) + [None] * (width - len(r))
        if not str(r[0] or "").strip():
            continue
        out.append([CONVERT[c["type"]](r[i]) for i, c in enumerate(cols)])
    return out


def parse(path, progress=lambda msg: None):
    wb = load_workbook(str(path))
    names = wb.sheet_names
    for need in ("Full Ageing", "Account Analyst List"):
        if need not in names:
            raise ValueError(f"Sheet '{need}' not found in the Ageing file.")
    progress("Reading Full Ageing")
    full = _table(wb.get_sheet_by_name("Full Ageing").to_python(), "ageing_full", "Acct", 21)
    progress("Reading Account Analyst List")
    analyst = _table(wb.get_sheet_by_name("Account Analyst List").to_python(), "account_analyst", "Division", 8)
    return {"ageing_full": full, "account_analyst": analyst}
