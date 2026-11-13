"""A/R Inquiry PDF. The PDF is read in the browser with pdf.js and the Report2Excel parser (static/parse_report.js);
the browser posts the parsed objects here and this module maps them onto the SQLite tables."""
from .common import to_date, to_num, to_text
from .specs import TABLES

# spec key -> key used by parse_report.js
ALIASES = {
    "ar_detail": {"cust_code": "custCode", "pd_date": "pdDate", "check_amt": "checkAmt", "check_date": "checkDate"},
    "ar_aging": {"cust_code": "custCode"},
    "ar_customers": {"ytd_ret": "ytdRet"},
}
SOURCE_KEY = {"ar_detail": "detail", "ar_aging": "aging", "ar_customers": "customers"}
CONVERT = {"s": to_text, "n": to_num, "d": to_date}


def rows_from_payload(payload: dict) -> dict:
    out = {}
    for tid, src_key in SOURCE_KEY.items():
        cols, alias, rows = TABLES[tid]["cols"], ALIASES.get(tid, {}), payload.get(src_key) or []
        out[tid] = [[CONVERT[c["type"]](r.get(alias.get(c["key"], c["key"]))) for c in cols] for r in rows]
    if not out["ar_detail"]:
        raise ValueError("No A/R detail lines found in this PDF.")
    return out
