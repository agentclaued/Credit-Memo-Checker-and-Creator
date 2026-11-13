"""Reason code list (ReasonCodeList.xlsx / .xls / .csv): two columns, Reason Code and Description."""
import csv
import re

from python_calamine import load_workbook

from .common import to_text


def fix_code(v):
    """Reason codes are stored in UPPER CASE."""
    return str(v or "").strip().upper()


def fix_desc(v):
    """Descriptions keep their own capitals (the batch REASON sheet is the reference); only extra spaces are removed."""
    return re.sub(r"\s+", " ", str(v or "")).strip() or None


def _rows(path):
    if str(path).lower().endswith(".csv"):
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            yield from csv.reader(f)
    else:
        wb = load_workbook(str(path))
        yield from wb.get_sheet_by_name(wb.sheet_names[0]).to_python()


def parse(path, progress=lambda msg: None):
    rows = [r for r in _rows(path) if r]
    if not rows or "code" not in str(rows[0][0]).lower():
        raise ValueError("Expected a reason code list: first column 'Reason Code', second column 'Description'.")
    out, seen = [], set()
    for r in rows[1:]:
        code = fix_code(to_text(r[0]))
        if not code or code in seen:
            continue
        seen.add(code)
        out.append([code, fix_desc(to_text(r[1])) if len(r) > 1 else None])
    if not out:
        raise ValueError("No reason codes found.")
    return {"reason_list": out}
