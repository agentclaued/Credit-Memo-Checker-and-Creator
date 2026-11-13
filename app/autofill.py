"""Fills memo details from the imported data, for the DB# / IN# on each Batch row (Ageing, Account list, Report 2, Sale Report)."""
import re

from . import recheck


def _name(v) -> str:
    """'J35 JACK/AARON' -> 'JACK/AARON', 'A4 NET 60 DAYS' -> 'NET 60 DAYS' (the leading code is not written on the memo)."""
    return re.sub(r"^\S*\d\S*\s+", "", str(v or "").strip())


def lookup(refs: list, division: str = "") -> dict:
    """{ref as typed: {found, acct, name, div, claim, total, salesman, terms, shipAcct}} for each DB#/IN# of the Batch sheet."""
    ref = recheck._ref_data()
    out = {}
    for raw in refs:
        key = recheck._norm(raw)
        rows = ref["ageing"].get(key) or []
        if not key or not rows:
            out[raw] = {"found": False}
            continue
        a = rows[0]
        acct = recheck._norm(a["acct"])
        divs = sorted(ref["ageing_div"].get(acct, ()))
        div = division.upper() if division.upper() in divs else (divs[0] if divs else division.upper())
        cust = recheck._norm(div) + acct
        terms_slm = sorted(ref["cust"].get(cust, ()))
        an = ref["analyst"].get(acct) or {}
        terms = _name(terms_slm[0][0] if terms_slm else an.get("terms"))
        slm = _name(terms_slm[0][1] if terms_slm else an.get("slm"))
        ships = ref["sale_ship"].get(cust) or ref["ar_ship"].get(cust) or set()
        out[raw] = {"found": True, "acct": acct, "name": (a["name"] or "").strip(), "div": div, "claim": (a["claim_no"] or "").strip(),
                    "total": a["total"], "salesman": slm, "terms": terms, "shipAcct": next(iter(ships)) if len(ships) == 1 else ""}
    return out
