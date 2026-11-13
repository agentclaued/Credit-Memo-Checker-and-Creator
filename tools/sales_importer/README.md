# Sales Upload: raw file to table

`sales_import.py` is the dashboard's Sales upload (the "Orders by Item" export) as one standalone file.
It is a copy of the logic in `app/importers.py` and the `/upload` route in `app/main.py`. It does not
change the dashboard and does not touch `app/reporting.db` unless you point `--db` at it.

## Setup (once)

```
pip install -r requirements.txt
```

## Run

```
python sales_import.py "C:\path\Orders_by_Item.xlsx" --region USA
python sales_import.py "C:\path\Orders_by_Item.csv" --region Canada --db my_test.db --user Maria
```

| Option | Meaning |
|---|---|
| `file` | CSV, XLS, XLSX, XLSM, XLSB, ODS, HTML or ZIP export |
| `--region` | `USA` or `Canada` (required, same as the Import center dropdown) |
| `--db` | SQLite file to write. Default `sales_demo.db`, created if missing |
| `--user` | Stored in `uploaded_by` |
| `--company-id` | Stored in `company_id` (default 1) |

## What happens, step by step

1. **Read** (`read_sheets`): every worksheet is read; empty sheets are ignored.
2. **Clean** (`clean_sales_rows`):
   - finds the header row (needs an Amount column plus Invoice/Order/PO column);
   - reads the Bill-To code (like `WA9085`) from cell A3 and adds it as **EM Acct**;
   - copies each item-description row onto its orders as **Item Name**, then drops the description row;
   - drops blank rows and repeated page headers.
3. **Validate and save** (`import_sales`): each row needs an invoice, order or PO number. Amounts like
   `$1,200.00` or `(50.00)` are converted to numbers. A bad row is skipped and reported, and the rest still load.
4. **Table**: rows go into `records` with `record_type = 'invoices'`.
   - Row key is `invoice-or-order|microfilm|invoices`. Uploading the same sales again **updates** the
     existing rows instead of duplicating them. A repeat inside one file gets `|2`, `|3`.
   - `source_file` stores the file name and sheet for every row.
   - `raw_data` stores every original column as JSON, dates shown as MM/DD/YYYY.
   - The whole file is saved in one transaction, so a crash saves nothing.

## Table: `records`

`company_id, record_type, record_key, invoice_number, po_number, customer, record_date, amount,
paid_amount, status, deduction_code, acct, region, uploaded_by, source_file, created_at, raw_data, hidden`

## Using it in your own code

```python
import sqlite3
from sales_import import SCHEMA, read_sheets, clean_sales_rows, import_sales

conn = sqlite3.connect("my.db"); conn.execute(SCHEMA)
for sheet, rows in read_sheets(open("file.xlsx", "rb").read(), "file.xlsx"):
    rows, removed = clean_sales_rows(rows)
    result = import_sales(conn, rows, f"file.xlsx [{sheet}]", region="USA")
    print(result.added, result.updated, result.errors)
conn.commit()
```

## Notes

- To check an EM Acct against the region you chose, pass `account_regions={"WA9085": "USA"}` to
  `import_sales`. The dashboard takes this from its Aging data; the standalone copy skips it by default.
- To load into the live dashboard, use the Import center. It also bumps the company data version so
  cached reports refresh, which this script does not do.
