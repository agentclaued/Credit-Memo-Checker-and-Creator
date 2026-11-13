"""Sale Report = Inquiry Report 6 "Orders by Item" (csv/xlsx).

Layout of the raw file: a title block (line 3 holds "Bill-To #=EMAM9286": Dv "EM" + account "AM9286"), then one block per item:
a header row, the order lines, and an item-total row that carries the item number, description and pack note.
Clean result: one row per order line, with the account (SaleReport) in front, the item Description and Pack copied from the
item-total row (matched on Item#), and Dv as in the file. Headers, item totals, grand totals and blank rows are dropped.
"""
import csv
import re

from python_calamine import load_workbook

from .common import CONVERT, to_text
from .specs import TABLES

ORDER_COLS = ["dv", "acct", "name", "dept", "order_no", "date", "tm", "slm", "store_no", "cust_sku", "item_no", "qty", "uom",
              "price", "amount", "currency", "status", "deliv", "cancel", "pt_no", "whs", "inv_no", "rec_no", "from_whse"]
BILL_TO = re.compile(r"Bill-To\s*#\s*=\s*([A-Za-z]{2})([A-Za-z0-9]+)")     # =EM + AM9286


def _rows_from_file(path):
    if str(path).lower().endswith(".csv"):
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            yield from csv.reader(f)
    else:
        wb = load_workbook(str(path))
        for name in wb.sheet_names:
            yield from wb.get_sheet_by_name(name).to_python()


def _cell(r, i):
    return str(r[i]).strip() if len(r) > i and r[i] is not None else ""


def _is_order_line(r):
    return len(r) >= 15 and _cell(r, 0) not in ("", "Dv") and bool(_cell(r, 4)) and _cell(r, 4) != "Order#"


def _is_item_total(r):
    """Dv, item number, description, pack note, and no order number."""
    return len(r) > 3 and _cell(r, 0) not in ("", "Dv") and bool(_cell(r, 1)) and not _cell(r, 4) and bool(_cell(r, 2))


def parse(path, progress=lambda msg: None):
    spec = {c["key"]: c for c in TABLES["orders"]["cols"]}
    out, group, account = [], [], None
    for n, r in enumerate(_rows_from_file(path)):
        r = [("" if v is None else v) for v in r]
        if account is None:
            m = BILL_TO.search(" ".join(str(v) for v in r))
            if m:
                account = m.group(2).upper()
        if _is_order_line(r):
            group.append(r)
        elif group and _is_item_total(r):
            item, desc, pack = _cell(r, 1), to_text(r[2]), to_text(r[3])
            for g in group:
                same = _cell(g, 10) == item
                out.append(_build(g, account, desc if same else None, pack if same else None, spec))
            group = []
            if len(out) % 5000 < 50:
                progress(f"Read {len(out):,} order lines")
    for g in group:       # lines with no closing item-total row
        out.append(_build(g, account, None, None, spec))
    if not out:
        raise ValueError("No order lines found (expected an 'Orders by Item' sale report).")
    if not account:
        raise ValueError("Account not found: line 3 should read like 'Bill-To #=EMAM9286'.")
    return {"orders": out}


def _build(g, account, item_name, pack, spec):
    g = list(g) + [""] * (len(ORDER_COLS) - len(g))
    vals = {k: CONVERT[spec[k]["type"]](g[i]) for i, k in enumerate(ORDER_COLS)}
    vals.update(sale_report=account, item_name=item_name, pack=pack,
                dv_sale_report=f"{vals['dv'] or ''}{account or ''}" or None)
    return [vals.get(c["key"]) for c in TABLES["orders"]["cols"] if not c.get("virt")]
