"""Excel export of a Create-batch draft: a copy of the strict template, filled only in the template's own cells.

Only the cells and formulas of templates_xlsx/Credit Memo Batch TEMPLATE.xlsx are used (see template_rules.py);
the template formulas (VLOOKUP, price x qty, SUBTOTAL, TOTAL, Batch total) are kept and Excel calculates them.
"""
import io
from copy import copy
from datetime import date, datetime

import openpyxl

from . import template_rules as T


def _num(v):
    try:
        return float(v) if v not in ("", None) else None
    except (TypeError, ValueError):
        return None


def _date(v):
    if isinstance(v, (date, datetime)):
        return v
    for fmt, n in (("%Y-%m-%d", 10), ("%m/%d/%Y", None)):
        try:
            return datetime.strptime(str(v).strip()[:n], fmt)
        except ValueError:
            continue
    return None


def _money(v):
    import re
    t = re.sub(r"[^0-9.\-]", "", str(v or ""))
    try:
        return float(t) if t not in ("", "-", ".") else None
    except ValueError:
        return None


def _total(m) -> float:
    sub = sum((_num(l.get("price")) or 0) * (_num(l.get("qty")) or 0) for l in m.get("lines", []))
    return round(sub + sum(_num(m.get(k)) or 0 for k in ("gst", "freight", "adj")), 2)


def _grow_batch(batch, extra: int) -> None:
    """More than the template's 20 Batch rows: add numbered rows in the style of the last row and move the total row down."""
    if extra <= 0:
        return
    last, tot = T.BATCH_LAST_ROW, T.BATCH_LAST_ROW + 1
    for c in range(1, 11):
        src, t = batch.cell(last, c), batch.cell(tot, c)
        for k in range(1, extra + 1):
            d = batch.cell(tot + k - 1, c)
            d._style = copy(src._style)
        nt = batch.cell(tot + extra, c)
        nt._style, nt.value = copy(t._style), t.value
    for k in range(1, extra + 1):
        batch.cell(last + k, 1, (last - T.BATCH_FIRST_ROW + 1) + k)
    batch.cell(tot + extra, 4, f"=SUM(D{T.BATCH_FIRST_ROW}:D{last + extra})")
    total_height = batch.row_dimensions[tot].height
    for k in range(extra):
        batch.row_dimensions[tot + k].height = batch.row_dimensions[last].height
    batch.row_dimensions[tot + extra].height = total_height


def build(data: dict, template_path, price_rows: list, reasons: dict) -> io.BytesIO:
    """data = {"batch": {no, date, return, division}, "memos": [{...}]}; reasons = {code: description}."""
    memos = data.get("memos") or []
    rows_in = data.get("batchRows")
    if not memos and not any(any(str(v).strip() for v in r.values()) for r in rows_in or []):
        raise ValueError("Add at least one memo or fill the Batch sheet first.")
    wb = openpyxl.load_workbook(template_path)
    batch, first = wb["Batch"], wb["1"]
    used_rows = max(len(memos), max((i + 1 for i, r in enumerate(rows_in or []) if any(str(v).strip() for v in r.values())), default=0))
    _grow_batch(batch, used_rows - (T.BATCH_LAST_ROW - T.BATCH_FIRST_ROW + 1))

    pl = wb["PRICE LIST"]
    typed = [r for r in (data.get("priceList") or []) if str(r.get("item", "")).strip() or str(r.get("desc", "")).strip()]
    if typed:                                  # the PRICE LIST sheet as edited on screen (typed or pasted from Excel)
        for i, r in enumerate(typed, 2):
            pl.cell(i, 1, str(r.get("item", "")).strip() or None)
            pl.cell(i, 2, str(r.get("desc", "")).strip() or None)
    elif pl.max_row < 2 and price_rows:          # the template's PRICE LIST is empty: fill it so the VLOOKUP formulas can find the descriptions
        for i, row in enumerate(price_rows, 1):
            for j, v in enumerate(row[:2], 1):
                pl.cell(i, j, v)

    b = data.get("batch") or {}
    for key, cell in T.BATCH_HEADER_CELLS.items():
        v = _date(b.get(key)) if key == "date" else (b.get(key) or None)
        if v is not None:
            batch[cell] = v

    sheets = []
    if rows_in:                                    # the Batch sheet as edited on screen (typed or pasted from Excel)
        for i, row in enumerate(rows_in):
            r = T.BATCH_FIRST_ROW + i
            if not any(str(v).strip() for v in row.values()):
                continue
            batch.cell(r, 2, (row.get("company") or "").strip() or None)
            batch.cell(r, 3, _date(row.get("date")))
            batch.cell(r, 4, _money(row.get("amount")))
            batch.cell(r, 5, (row.get("db") or "").strip() or None)
            batch.cell(r, 6, (row.get("claim") or "").strip() or None)
            batch.cell(r, 7, (row.get("rt") or "").strip() or None)
            for col, key in ((8, "rtAmt"), (9, "disc"), (10, "over")):
                batch.cell(r, col, _money(row.get(key)))

    for n, m in enumerate(memos or [{}], 1):
        ws = first if n == 1 else wb.copy_worksheet(first)
        ws.title = str(n)
        if ws is not first:          # copy_worksheet drops the view: keep memo 1's zoom and gridlines
            ws.sheet_view.zoomScale = first.sheet_view.zoomScale
            ws.sheet_view.zoomScaleNormal = first.sheet_view.zoomScaleNormal
            ws.sheet_view.showGridLines = first.sheet_view.showGridLines
        sheets.append(ws)
        code = (m.get("code") or "").strip()
        reason = reasons.get(code, "")
        ws["B4"], ws["B6"] = m.get("soldTo") or None, m.get("soldAcct") or None
        ws["G4"], ws["G6"] = m.get("shipTo") or None, m.get("shipAcct") or None
        ws["H2"] = _date(m.get("date"))
        ws["A9"], ws["C9"] = reason or None, code or None
        ws["E9"], ws["F9"], ws["H9"] = m.get("salesman") or None, m.get("terms") or None, m.get("via") or None
        claim = (m.get("claim") or "").strip()
        ws["E10"] = (f"{reason} : {claim}" if claim else reason) or None
        for i, l in enumerate((m.get("lines") or [])[: T.MEMO_LAST_LINE - T.MEMO_FIRST_LINE + 1]):
            r = T.MEMO_FIRST_LINE + i
            sku = (l.get("sku") or "").strip()
            if sku:
                ws[f"A{r}"] = sku
            if _num(l.get("qty")) is not None:
                ws[f"D{r}"] = _num(l["qty"])
            if _num(l.get("price")) is not None:
                ws[f"G{r}"] = _num(l["price"])
        for row, key in ((35, "gst"), (36, "freight"), (37, "adj")):
            if _num(m.get(key)) is not None:
                ws[f"I{row}"] = _num(m[key])
        ws["A40"] = m.get("db") or None

        if rows_in or not m:
            continue
        r = T.BATCH_FIRST_ROW + n - 1
        batch.cell(r, 2, m.get("soldTo") or None)
        batch.cell(r, 3, _date(m.get("date")))
        batch.cell(r, 4, _total(m))
        batch.cell(r, 5, m.get("db") or None)
        batch.cell(r, 6, claim or None)
        for col, key in ((7, "rt"), (8, "rtAmt"), (9, "disc"), (10, "over")):
            v = m.get(key)
            batch.cell(r, col, _num(v) if col > 7 and _num(v) is not None else (v or None) if col == 7 else None)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
