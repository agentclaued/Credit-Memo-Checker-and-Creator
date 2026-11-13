"""Credit Memo Batch workbook (the template, or a filled batch to validate).

Reads only cells that exist in the template: REASON, PRICE LIST, Batch and the numbered memo sheets.
"""
from python_calamine import load_workbook

from .common import to_date, to_num, to_text


def _cell(rows, r, c):
    try:
        return rows[r][c]
    except IndexError:
        return None


def parse(path, progress=lambda msg: None):
    wb = load_workbook(str(path))
    names = wb.sheet_names
    for need in ("REASON", "PRICE LIST", "Batch"):
        if need not in names:
            raise ValueError(f"Sheet '{need}' not found. This is not a Credit Memo Batch file.")
    out = {"price_list": [], "reason_codes": [], "batch_rows": [], "memo_headers": [], "memo_lines": []}

    for r in wb.get_sheet_by_name("PRICE LIST").to_python()[1:]:
        if to_text(_cell([r], 0, 0)):
            out["price_list"].append([to_text(r[0]), to_text(r[1] if len(r) > 1 else None)])
    for r in wb.get_sheet_by_name("REASON").to_python()[2:]:
        if to_text(_cell([r], 0, 0)):
            out["reason_codes"].append([to_text(r[0]), to_text(r[1] if len(r) > 1 else None)])

    batch = wb.get_sheet_by_name("Batch").to_python()
    start = next((i for i, r in enumerate(batch) if r and str(r[0]).strip() == "#"), None)
    if start is None:
        raise ValueError("Batch header row ('#', COMPANY, DATE...) not found.")
    for r in batch[start + 1:]:
        r = list(r) + [None] * (10 - len(r))
        if not isinstance(r[0], (int, float)) or not to_text(r[1]) and to_num(r[3]) is None:
            continue
        out["batch_rows"].append([to_num(r[0]), to_text(r[1]), to_date(r[2]), to_num(r[3]), to_text(r[4]), to_text(r[5]),
                                  to_text(r[6]), to_num(r[7]), to_num(r[8]), to_num(r[9])])

    memos = sorted((n for n in names if n.strip().isdigit()), key=int)
    for n in memos:
        rows = wb.get_sheet_by_name(n).to_python()
        no = float(n)
        sub_row = next((i for i, r in enumerate(rows) if str(_cell(rows, i, 4) or "").strip() == "SUBTOTAL"), None)
        if sub_row is None:
            continue
        tot_row = next((i for i in range(sub_row, len(rows)) if str(_cell(rows, i, 6) or "").strip() == "TOTAL"), None)
        db_row = next((i for i, r in enumerate(rows) if str(_cell(rows, i, 0) or "").strip() == "DEBIT MEMO NOS."), None)
        before = len(out["memo_lines"])
        for i in range(12, sub_row):                      # item rows start at row 13
            sku = to_text(_cell(rows, i, 0))
            if not sku:
                continue
            out["memo_lines"].append([no, sku, to_num(_cell(rows, i, 3)), to_text(_cell(rows, i, 4)),
                                      to_num(_cell(rows, i, 6)), to_num(_cell(rows, i, 8))])
        if not any(to_text(_cell(rows, r, c)) for r, c in ((3, 1), (3, 6), (8, 0))) and len(out["memo_lines"]) == before:
            continue   # empty form (the template's sample sheet): nothing to import
        out["memo_headers"].append([
            no, to_date(_cell(rows, 1, 7)), to_text(_cell(rows, 3, 1)), to_text(_cell(rows, 5, 1)),
            to_text(_cell(rows, 3, 6)), to_text(_cell(rows, 5, 6)), to_text(_cell(rows, 8, 0)), to_text(_cell(rows, 8, 2)),
            to_text(_cell(rows, 8, 4)), to_text(_cell(rows, 8, 5)), to_text(_cell(rows, 8, 7)), to_text(_cell(rows, 9, 4)),
            to_num(_cell(rows, sub_row, 8)), to_num(_cell(rows, sub_row + 1, 8)), to_num(_cell(rows, sub_row + 2, 8)),
            to_num(_cell(rows, sub_row + 3, 8)), to_num(_cell(rows, tot_row, 8)) if tot_row is not None else None,
            to_text(_cell(rows, db_row + 1, 0)) if db_row is not None else None])
        progress(f"Read memo sheet {n}")
    return out
