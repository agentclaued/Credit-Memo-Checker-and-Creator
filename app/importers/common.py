"""Value converters shared by the importers."""
import re
from datetime import date, datetime

_DATE_FORMATS = ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S")


def to_text(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s or None


def to_num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("$", "").replace(",", "")
    neg = s.startswith("(") and s.endswith(")") or s.endswith("-")
    s = s.strip("()-")
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if neg else n


def to_date(v):
    """Return YYYY-MM-DD or None."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    for f in _DATE_FORMATS:
        try:
            return datetime.strptime(s, f).date().isoformat()
        except ValueError:
            pass
    return None


CONVERT = {"s": to_text, "n": to_num, "d": to_date}


def money(s):
    """pdf-style money token: 1,234.56 or 1,234.56- (trailing minus = negative)."""
    if not re.fullmatch(r"[\d,]*\.\d\d-?", s or ""):
        return None
    neg = s.endswith("-")
    n = float(s.replace(",", "").replace("-", ""))
    return -n if neg else n
