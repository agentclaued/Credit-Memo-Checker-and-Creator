"""Table registry. One entry per imported table: it defines the SQLite table AND the columns the viewer shows.

Column = (key, label, type, options). type: s = text, n = number, d = date (stored as YYYY-MM-DD).
options: dec (decimals), cur (show $), sum (total row sums the filtered rows), plain (number without separators).
"""

C = lambda key, label, typ="s", **o: {"key": key, "label": label, "type": typ, **o}
MONEY = dict(dec=2, cur=True)

TABLES = {
    # ---------------------------------------------------------------- ageing xlsx
    "ageing_full": {"kind": "ageing", "title": "Full Ageing", "cols": [
        C("acct", "Acct"), C("div", "Div"), C("name", "Name"), C("ref_no", "Ref#"), C("ref_type", "Ref Type"),
        C("date", "Date", "d"), C("due_date", "Due Date", "d"), C("order_no", "Order#"),
        C("total", "Total", "n", sum=True, **MONEY), C("current", "Current", "n", **MONEY),
        C("d1_30", "1-30 Days", "n", **MONEY), C("d31_60", "31-60 Days", "n", **MONEY),
        C("d61_90", "61-90 Days", "n", **MONEY), C("over90", "Over 90 Days", "n", **MONEY),
        C("claim_no", "Claim#"), C("comment", "Comment"), C("assigned", "Assigned Person"), C("status", "Status"),
        C("check_no", "Check#"), C("check_amt", "Check Amt", "n", **MONEY), C("check_date", "Check Date", "d")]},
    "account_analyst": {"kind": "ageing", "title": "Account Analyst List", "cols": [
        C("division", "Division"), C("acct", "Acct"), C("company", "Company Name"), C("analyst", "Analyst"),
        C("terms", "Terms"), C("slm", "SLM"), C("vendor", "Vendor#"), C("assigned", "Assigned Person")]},
    # ---------------------------------------------------------------- orders csv
    "orders": {"kind": "orders", "title": "Sale Report", "cols": [
        C("sale_report", "SaleReport"),
        C("dv", "Dv"), C("acct", "Acct ShipCode"), C("dv_sale_report", "Dv+SaleReport"), C("name", "Name"), C("dept", "Dept"), C("order_no", "Order#"),
        C("date", "Date", "d"), C("tm", "Tm"), C("slm", "Slm"), C("store_no", "Store#"), C("cust_sku", "CustSKU"),
        C("item_no", "Item#"), C("item_name", "Description"), C("pack", "Pack"), C("qty", "Qty", "n", dec=0, sum=True), C("uom", "UOM"),
        C("price", "Price", "n", **MONEY), C("amount", "Amount", "n", sum=True, **MONEY), C("currency", "Currency"),
        C("status", "Status"), C("deliv", "Deliv", "d"), C("cancel", "Cancel", "d"), C("pt_no", "P.T.#"),
        C("whs", "Whs"), C("inv_no", "Inv#"), C("rec_no", "Rec#"), C("from_whse", "From Whse"),
        C("source_file", "Source File", virt=True)]},   # virt: shown from the built-in source_file column, not stored twice
    # ---------------------------------------------------------------- A/R inquiry PDF (same tables as the Report2Excel tool)
    "ar_detail": {"kind": "ar_pdf", "title": "A/R Detail", "cols": [
        C("page", "Pg#", "n", plain=True), C("acct", "Acct"), C("dv", "Div"), C("customer", "Name"), C("ref", "Ref#"),
        C("type", "Ref Type"), C("date", "Date", "d"), C("due", "Due Date", "d"), C("order", "Order#"), C("past", "Past"),
        C("amount", "Amount", "n", sum=True, **MONEY), C("claim", "Claim#"), C("ship", "Ship Code"), C("cust_code", "Cust Code"),
        C("terms", "Terms"), C("slm", "SLM"), C("vendor", "Vendor#"), C("pd_date", "Pd. Date", "d"), C("check", "Check#"),
        C("check_amt", "Check Amt", "n", **MONEY), C("check_date", "Check Date", "d")]},
    "ar_aging": {"kind": "ar_pdf", "title": "Aging Summary", "cols": [
        C("page", "Pg#", "n", plain=True), C("scope", "Account"), C("cust_code", "Cust Code"), C("name", "Customer"),
        C("type", "Type"), C("total", "Total", "n", **MONEY), C("current", "Current", "n", **MONEY),
        C("d1_30", "1 - 30 Days", "n", **MONEY), C("d31_60", "31 - 60 Days", "n", **MONEY),
        C("d61_90", "61 - 90 Days", "n", **MONEY), C("over90", "Over 90 Days", "n", **MONEY)]},
    "ar_customers": {"kind": "ar_pdf", "title": "Accounts", "cols": [
        C("page", "Pg#", "n", plain=True), C("code", "Cust Code"), C("name", "Customer"), C("terms", "Terms"), C("slm", "SLM"),
        C("vendor", "Vendor#"), C("balance", "Balance", "n", sum=True, **MONEY), C("mtd", "M-T-D Sales", "n", **MONEY),
        C("ytd", "Y-T-D Sales", "n", **MONEY), C("ytd_ret", "Y-T-D Retns", "n", **MONEY)]},
    # ---------------------------------------------------------------- reason code list (editable)
    "reason_list": {"kind": "reasons", "title": "Reason codes", "cols": [C("code", "Reason Code"), C("description", "Description")]},
    # ---------------------------------------------------------------- credit memo batch files (template and file to validate)
    "price_list": {"kind": "batch", "title": "PRICE LIST", "cols": [C("item", "Item #"), C("description", "Description")]},
    "reason_codes": {"kind": "batch", "title": "REASON", "cols": [C("code", "Code"), C("description", "Description")]},
    "batch_rows": {"kind": "batch", "title": "Batch", "cols": [
        C("no", "#", "n", plain=True), C("company", "Company"), C("date", "Date", "d"), C("amount", "Amount", "n", sum=True, **MONEY),
        C("db_no", "DB#"), C("claim_no", "Claim #"), C("rt_no", "RT #"), C("rt_amount", "RT Amount", "n", **MONEY),
        C("discrepancy", "Discrepancy", "n", **MONEY), C("over_under", "Credit Over / Under", "n", **MONEY)]},
    "memo_headers": {"kind": "batch", "title": "Memo forms", "cols": [
        C("memo", "Memo", "n", plain=True), C("date", "Date", "d"), C("sold_to", "Sold To"), C("sold_acct", "Sold Acct"),
        C("ship_to", "Ship To"), C("ship_acct", "Ship Acct"), C("reason", "Reason"), C("reason_code", "Reason Code"),
        C("salesman", "Salesman"), C("terms", "Terms"), C("via", "Shipped Via"), C("claim", "Customer Claim #"),
        C("subtotal", "Subtotal", "n", **MONEY), C("gst", "GST/HST", "n", **MONEY), C("freight", "Freight/Handling", "n", **MONEY),
        C("adjustment", "Adjustment", "n", **MONEY), C("total", "Total", "n", sum=True, **MONEY), C("db_no", "Debit Memo Nos.")]},
    "memo_lines": {"kind": "batch", "title": "Memo lines", "cols": [
        C("memo", "Memo", "n", plain=True), C("style", "Style No."), C("qty", "Quantity PCS", "n", dec=0, sum=True),
        C("description", "Description"), C("price", "Price", "n", **MONEY), C("amount", "Amount", "n", sum=True, **MONEY)]},
}

# Which tables each uploaded file kind produces (viewer tab order).
KIND_TABLES = {
    "ageing": ["ageing_full", "account_analyst"],
    "orders": ["orders"],
    "ar_pdf": ["ar_detail", "ar_aging", "ar_customers"],
    "reasons": ["reason_list"],
    "template": ["price_list", "reason_codes", "batch_rows"],
    "validate": ["batch_rows", "memo_headers", "memo_lines", "price_list", "reason_codes"],
}

SQL_TYPE = {"s": "TEXT", "n": "REAL", "d": "TEXT"}


def schema_sql() -> str:
    out = []
    for tid, t in TABLES.items():
        cols = ",\n    ".join(f"\"{c['key']}\" {SQL_TYPE[c['type']]}" for c in t["cols"] if not c.get("virt"))
        out.append(f"CREATE TABLE IF NOT EXISTS t_{tid} (\n    id INTEGER PRIMARY KEY,\n    source_file_id INTEGER NOT NULL,\n"
                   f"    source_file TEXT NOT NULL,\n    row_no INTEGER NOT NULL,\n    {cols}\n);\n"
                   f"CREATE INDEX IF NOT EXISTS ix_t_{tid}_file ON t_{tid}(source_file_id);\n"
                   f"CREATE INDEX IF NOT EXISTS ix_t_{tid}_file_row ON t_{tid}(source_file_id, row_no);")
    return "\n".join(out)
