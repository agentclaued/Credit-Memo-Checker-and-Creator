"""Standalone Sales upload: raw file (Orders by Item export) -> cleaned rows -> SQLite `records` table.

This is the same logic the dashboard uses (app/importers.py, `invoices` kind), copied into one file
that runs on its own. Nothing here touches app/ or the live database unless you point --db at it.

Usage:
    python sales_import.py <file> --region USA|Canada [--db sales_demo.db] [--user name] [--company-id 1]
"""
import argparse
import csv
import io
import json
import re
import sqlite3
import sys
import zipfile
from datetime import date, datetime
from html.parser import HTMLParser
from pathlib import Path

from openpyxl import load_workbook
from python_calamine import load_workbook as load_calamine

SUPPORTED_EXTENSIONS = {".csv", ".htm", ".html", ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".zip"}
DATE_FORMATS = ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y")

AMOUNT_HEADERS = {"amount", "invoice amount", "invoice amount ($)", "invoice amount($)", "total amount", "sales amount"}
ID_HEADERS = {"invoice number", "invoice #", "invoice no", "order#", "order #", "order number",
              "po number", "po #", "purchase order"}
INVOICE_KEYS = ("invoice number", "invoice #", "invoice no")
ORDER_KEYS = ("order#", "order #", "order number")
PO_KEYS = ("po number", "po #", "purchase order")
AMOUNT_KEYS = ("amount", "invoice amount", "invoice amount ($)", "invoice amount($)", "total amount", "sales amount")
REGIONS = ("USA", "Canada")

# Same shape as the dashboard's `records` table (app/company_schema.py).
SCHEMA = """
CREATE TABLE IF NOT EXISTS records (
    id INTEGER PRIMARY KEY,
    company_id INTEGER NOT NULL,
    record_type TEXT NOT NULL,
    record_key TEXT NOT NULL,
    invoice_number TEXT,
    po_number TEXT,
    customer TEXT,
    record_date TEXT,
    amount REAL NOT NULL DEFAULT 0,
    paid_amount REAL NOT NULL DEFAULT 0,
    status TEXT,
    deduction_code TEXT,
    acct TEXT NOT NULL DEFAULT '',
    region TEXT NOT NULL DEFAULT '',
    uploaded_by TEXT NOT NULL DEFAULT '',
    source_file TEXT NOT NULL,
    created_at TEXT NOT NULL,
    raw_data TEXT NOT NULL DEFAULT '{}',
    hidden INTEGER NOT NULL DEFAULT 0,
    UNIQUE (company_id, record_key)
)"""


# ---------------------------------------------------------------- cell values

def header(value) -> str:
    """Lower-case a header and collapse whitespace so column lookups are forgiving."""
    return " ".join(str(value or "").replace("\xa0", " ").strip().lower().split())


def to_number(value) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("$", "").replace(",", "")
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    try:
        return float(text) if text else 0.0
    except ValueError:
        raise ValueError(f"'{value}' is not a valid amount") from None


def to_iso_date(value):
    """YYYY-MM-DD when the value is a recognizable date, otherwise the original text."""
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return text


_ISO_DATE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T]\d{2}:\d{2}:\d{2})?$")
_SHORT_US_DATE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{2})$")


def to_mmddyyyy(value):
    """Format a date-like value as MM/DD/YYYY; anything else passes through unchanged."""
    if isinstance(value, (datetime, date)):
        return value.strftime("%m/%d/%Y")
    if not isinstance(value, str):
        return value
    text = value.strip()
    match = _ISO_DATE.match(text)
    if match:
        year, month, day = match.groups()
        return f"{month}/{day}/{year}"
    match = _SHORT_US_DATE.match(text)
    if match:
        month, day, year = match.groups()
        return f"{int(month):02d}/{int(day):02d}/20{year}"
    return value


# ---------------------------------------------------------------- 1. read the raw file

class _HtmlTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows, self._row, self._cell = [], None, None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in {"td", "th"} and self._row is not None:
            self._cell = []

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag):
        if tag in {"td", "th"} and self._cell is not None:
            self._row.append("".join(self._cell).strip())
            self._cell = None
        elif tag == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None


def _html_rows(content: bytes) -> list[list]:
    parser = _HtmlTableParser()
    parser.feed(content.decode("utf-8-sig", errors="replace"))
    return parser.rows


def _is_html(content: bytes) -> bool:
    probe = content.lstrip().lstrip(b"\xef\xbb\xbf").lower()
    return probe.startswith((b"<html", b"<!doctype html"))


def _single_sheet(content: bytes, filename: str) -> list[list]:
    ext = Path(filename).suffix.lower()
    if ext == ".csv":
        return list(csv.reader(io.StringIO(content.decode("utf-8-sig"))))
    if ext == ".zip":
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            sheet = next((n for n in archive.namelist()
                          if n.lower().endswith((".htm", ".html")) and "sheet" in Path(n).name.lower()), None)
            if not sheet:
                raise ValueError("ZIP must include the exported sheet HTML file, such as Check_004218861_files/sheet001.htm.")
            return _html_rows(archive.read(sheet))
    rows = _html_rows(content)
    if not rows:
        raise ValueError("This file is only an Excel HTML frameset. Upload its sheet001.htm file, "
                         "export it as CSV/XLSX, or upload a ZIP with the complete *_files folder.")
    return rows


def read_sheets(content: bytes, filename: str) -> list[tuple[str, list[list]]]:
    """Return (sheet name, rows) for every worksheet in the upload."""
    ext = Path(filename).suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError("unsupported file type")
    if ext in {".xlsx", ".xlsm"}:
        book = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        sheets = []
        for ws in book.worksheets:
            ws.reset_dimensions()  # some exports declare a one-cell sheet; read what is really there
            sheets.append((ws.title, [list(r) for r in ws.iter_rows(values_only=True)]))
        return sheets
    if ext in {".xls", ".xlsb", ".ods"} and not _is_html(content):
        book = load_calamine(io.BytesIO(content))
        return [(name, book.get_sheet_by_name(name).to_python()) for name in book.sheet_names]
    return [("Sheet1", _single_sheet(content, filename))]


# ---------------------------------------------------------------- 2. clean the sales rows

def _row_is_empty(row) -> bool:
    return not any(cell not in (None, "") for cell in row)


def clean_sales_rows(rows: list[list]) -> tuple[list[list], int]:
    """Return (rows starting at the header, number of removed rows).

    "Orders by Item" exports list order rows followed by an item-description row. That description is
    copied onto its orders as Item Name, the Bill-To code from cell A3 is added as EM Acct, and the
    description rows themselves are dropped.
    """
    cleaned = [[c.replace("\xa0", " ").strip() if isinstance(c, str) else c for c in row] for row in rows]
    header_index = next((i for i, row in enumerate(cleaned)
                         if any(header(c) in AMOUNT_HEADERS for c in row) and any(header(c) in ID_HEADERS for c in row)), None)
    if header_index is None:
        return cleaned, 0

    source_header = cleaned[header_index]
    pos = {header(c): i for i, c in enumerate(source_header) if header(c)}
    body = cleaned[header_index + 1:]

    if not {"order#", "item#", "amount"} <= pos.keys():  # plain invoice list: just drop blank rows
        kept = [row for row in body if not _row_is_empty(row)]
        return [source_header, *kept], header_index + len(body) - len(kept)

    bill_to = str(cleaned[2][0] or "") if len(cleaned) > 2 and cleaned[2] else ""
    match = re.search(r"([A-Za-z]{2}\d{4})$", bill_to.rsplit("=", 1)[-1].strip())
    if not match:
        raise ValueError("Sales report cell A3 must end with a Bill-To code in the form AADDDD, such as WA9085.")
    em_acct = match.group(1).upper()
    order_col, item_col, width = pos["order#"], pos["item#"], len(source_header)

    def with_extra_columns(row, item_name=""):
        out = []
        for i, cell in enumerate(row[:width]):
            out.append(cell)
            if i == order_col:
                out.append(em_acct)
            if i == item_col:
                out.append(item_name)
        return out

    result = [with_extra_columns(source_header, "Item Name")]
    result[0][order_col + 1] = "EM Acct"
    removed, pending = header_index, []

    for row in body:
        row = list(row) + [""] * max(0, width - len(row))
        if _row_is_empty(row):
            removed += 1
            continue
        row_headers = {header(c) for c in row if header(c)}
        if {"order#", "item#"} <= row_headers:  # repeated page header
            result.extend(with_extra_columns(r) for r in pending)
            pending, removed = [], removed + 1
            continue
        if str(row[order_col] or "").strip():
            pending.append(row)
            continue
        item_name = str(row[pos["name"]] or "").strip() if "name" in pos else ""
        item_code = str(row[pos["acct"]] or "").strip() if "acct" in pos else ""
        if item_name and item_code:
            result.extend(with_extra_columns(r, item_name) for r in pending)
            pending = []
        removed += 1
    result.extend(with_extra_columns(r) for r in pending)
    return result, removed


# ---------------------------------------------------------------- 3. validate + write to the table

class ImportResult:
    def __init__(self):
        self.added, self.updated, self.errors = 0, 0, []


class _Row:
    """Look up cells in a data row by normalized header name."""

    def __init__(self, positions, values):
        self.positions, self.values = positions, values

    def get(self, *keys, default=""):
        for key in keys:
            index = self.positions.get(key)
            if index is not None and index < len(self.values):
                return self.values[index]
        return default

    def text(self, *keys) -> str:
        return str(self.get(*keys) or "").strip()

    def text_or_null(self, *keys):
        return self.text(*keys) or None


def _positions(header_row) -> dict[str, int]:
    return {header(cell): i for i, cell in enumerate(header_row) if header(cell)}


def _is_blank(row) -> bool:
    return not any(cell is not None and str(cell).strip() != "" for cell in row)


def _raw_json(header_row, row) -> str:
    """Keep every source column exactly as uploaded, in its original order (repeated names get "(2)")."""
    data, seen = {}, {}
    for i, name in enumerate(header_row):
        if i >= len(row):
            break
        label = str(name).strip() if name not in (None, "") else f"Column {i + 1}"
        seen[label] = seen.get(label, 0) + 1
        if seen[label] > 1:
            label = f"{label} ({seen[label]})"
        data[label] = to_mmddyyyy(row[i])
    return json.dumps(data, default=str)


def _upsert(conn: sqlite3.Connection, company_id: int, record: dict) -> bool:
    """Insert the record, or update the row with the same record_key. True when inserted."""
    exists = conn.execute("SELECT 1 FROM records WHERE company_id = ? AND record_key = ?",
                          (company_id, record["record_key"])).fetchone()
    columns = [c for c in record if c != "record_key"]
    if exists:
        assignments = ", ".join(f"{c} = ?" for c in columns)
        conn.execute(f"UPDATE records SET {assignments} WHERE company_id = ? AND record_key = ?",
                     [record[c] for c in columns] + [company_id, record["record_key"]])
        return False
    names = ["company_id", "record_key", *columns]
    conn.execute(f"INSERT INTO records ({', '.join(names)}) VALUES ({', '.join('?' * len(names))})",
                 [company_id, record["record_key"], *[record[c] for c in columns]])
    return True


def import_sales(conn: sqlite3.Connection, rows: list[list], source: str, region: str,
                 user: str = "", company_id: int = 1, account_regions: dict | None = None) -> ImportResult:
    """Validate cleaned sales rows and upsert them into `records` (record_type='invoices').

    `account_regions` maps an EM Acct to its region (e.g. {"WA9085": "USA"}). When given, a file whose
    own EM Acct belongs to a different region than the one you chose is rejected row by row.
    """
    if not rows:
        raise ValueError("The file is empty.")
    header_row, pos = rows[0], _positions(rows[0])
    has_invoice = any(k in pos for k in INVOICE_KEYS)
    if not any(k in pos for k in AMOUNT_KEYS):
        raise ValueError("Invoices/orders require an Amount column.")
    if not (has_invoice or any(k in pos for k in ORDER_KEYS + PO_KEYS)):
        raise ValueError("Invoices/orders require Invoice Number, Order#, or PO Number.")

    regions = account_regions or {}
    result, now = ImportResult(), datetime.now().isoformat(timespec="seconds")
    seen_keys: dict[str, int] = {}  # repeats of the same invoice in one file get "|2", "|3", ...
    for line, values in enumerate(rows[1:], start=2):
        if _is_blank(values):
            continue
        row = _Row(pos, values)
        try:
            invoice = row.text(*INVOICE_KEYS)
            order = str(row.get(*PO_KEYS) or row.get(*ORDER_KEYS) or "").strip()
            if not (invoice or order):
                raise ValueError("invoice number or order number is required")
            if "em acct" in pos:
                if not row.text("em acct"):
                    raise ValueError("EM Acct could not be read from cell A3")
                if not row.text("item name"):
                    raise ValueError("Item Name could not be found for this order")
            file_acct = row.text("em acct")
            if region and file_acct and regions.get(file_acct, region) != region:
                raise ValueError(f"file's own account {file_acct} is in {regions.get(file_acct)}, "
                                 f"not the {region} region you selected — upload rejected")

            amount = to_number(row.get(*AMOUNT_KEYS))
            paid = to_number(row.get("paid amount"))
            status = row.text("status") or ("Paid" if amount and paid >= amount else "Open")
            record_date = row.get("invoice date") or row.get("date")
            base_key = f"{invoice or order}|{row.text('microfilm number')}|invoices"
            seen_keys[base_key] = seen_keys.get(base_key, 0) + 1
            record = {
                "record_key": base_key if seen_keys[base_key] == 1 else f"{base_key}|{seen_keys[base_key]}",
                "record_type": "invoices",
                "invoice_number": invoice,
                "po_number": order,
                "customer": row.text_or_null("customer") or row.text_or_null("name"),
                "record_date": to_iso_date(record_date),
                "amount": amount,
                "paid_amount": paid,
                "status": status.title(),
                "deduction_code": row.text_or_null("deduction code"),
                "acct": "",
                "region": region,
                "uploaded_by": user,
                "source_file": source,  # every row keeps the file name it came from
                "created_at": now,
                "raw_data": _raw_json(header_row, values),
            }
            if _upsert(conn, company_id, record):
                result.added += 1
            else:
                result.updated += 1
        except Exception as exc:
            result.errors.append(f"Row {line}: {exc}")
    return result


# ---------------------------------------------------------------- run it from the command line

def main() -> int:
    parser = argparse.ArgumentParser(description="Import a raw Sales (Orders by Item) file into a SQLite records table.")
    parser.add_argument("file", help="CSV / XLS / XLSX / XLSM / XLSB / ODS / HTML / ZIP sales export")
    parser.add_argument("--region", required=True, choices=REGIONS, help="Region this file belongs to")
    parser.add_argument("--db", default="sales_demo.db", help="SQLite file to write (default: sales_demo.db, created if missing)")
    parser.add_argument("--user", default="", help="Name stored in uploaded_by")
    parser.add_argument("--company-id", type=int, default=1, help="company_id stored on every row (default 1)")
    args = parser.parse_args()

    path = Path(args.file)
    if not path.is_file():
        print(f"File not found: {path}")
        return 1

    # The whole file is written in ONE transaction: an unexpected crash leaves the database untouched.
    conn = sqlite3.connect(args.db)
    conn.execute(SCHEMA)
    total_added = total_updated = 0
    try:
        for sheet_name, rows in read_sheets(path.read_bytes(), path.name):
            if _is_blank_sheet(rows):
                continue
            rows, removed = clean_sales_rows(rows)
            source = f"{path.name} [{sheet_name}]"
            result = import_sales(conn, rows, source, args.region, args.user, args.company_id)
            total_added, total_updated = total_added + result.added, total_updated + result.updated
            print(f"{source}: {result.added} new, {result.updated} updated, {removed} non-order rows removed")
            for err in result.errors[:10]:
                print("  skipped ->", err)
        conn.commit()
    except Exception as exc:
        conn.rollback()
        print(f"Import failed, nothing was saved: {exc}")
        return 1
    finally:
        conn.close()
    print(f"Done. {total_added} new, {total_updated} updated -> {args.db}")
    return 0


def _is_blank_sheet(rows) -> bool:
    return not any(any(cell not in (None, "") for cell in row) for row in rows)


if __name__ == "__main__":
    sys.exit(main())
