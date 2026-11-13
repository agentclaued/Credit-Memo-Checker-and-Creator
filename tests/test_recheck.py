"""Regression tests for the recheck protocols (app/recheck.py, rules in PROTOCOLS.md).

A clean batch is built from the template and must give ZERO errors. Then one thing at a time is broken and the matching
protocol must report it. If a protocol is changed, these tests show at once if a valid batch would get a false error
or a real mistake would be missed.
"""
import datetime
import os
import tempfile

os.environ.setdefault("CREDIT_DB_PATH", os.path.join(tempfile.mkdtemp(), "test.sqlite"))

import openpyxl  # noqa: E402
import pytest  # noqa: E402

from app import recheck  # noqa: E402
from app.database import ROOT, init_schema, transaction  # noqa: E402

TEMPLATE = ROOT / "templates_xlsx" / "Credit Memo Batch TEMPLATE.xlsx"
ACCT, NAME, SHIP = "AM9286", "AMAZON.COM.CA, INC", "AXYEG2"
DAY = datetime.datetime(2026, 10, 16)
MEMOS = [  # (DB#, style, pcs, price, claim text suffix)
    ("DB#900001", "TST100", 10, 10.0, " : SPA-1"),
    ("DB#900002", "TST200", 5, 10.0, ""),
]
REASON_CODE = "DA"


def _template_reasons():
    ws = openpyxl.load_workbook(TEMPLATE)["REASON"]
    return {str(r[0]).strip(): str(r[1]).strip() for r in ws.iter_rows(min_row=3, values_only=True) if r and r[0]}


@pytest.fixture(scope="module", autouse=True)
def seeded():
    """Ageing, account list, sale report and reason codes in a scratch database."""
    init_schema()
    reasons = _template_reasons()
    with transaction() as c:
        def file(kind, name):
            return c.execute("INSERT INTO source_files (kind, filename, stored_path, size_bytes, status) VALUES (?,?,?,?, 'imported')",
                             (kind, name, name, 1)).lastrowid
        f = file("ageing", "TEST ageing.xlsx")
        for i, (db, style, pcs, price, claim) in enumerate(MEMOS, 1):
            c.execute("INSERT INTO t_ageing_full (source_file_id, source_file, row_no, acct, div, name, ref_no, total, claim_no) VALUES (?,?,?,?,?,?,?,?,?)",
                      (f, "TEST ageing.xlsx", i, ACCT, "EM", NAME, db, pcs * price, claim.replace(" : ", "")))
        c.execute("INSERT INTO t_account_analyst (source_file_id, source_file, row_no, acct, company, terms, slm) VALUES (?,?,?,?,?,?,?)",
                  (f, "TEST ageing.xlsx", 1, ACCT, NAME, "NET 60 DAYS", "JACK/AARON"))
        o = file("orders", "TEST inquiry-report6.csv")
        for i, (db, style, *_rest) in enumerate(MEMOS, 1):
            c.execute("INSERT INTO t_orders (source_file_id, source_file, row_no, dv, sale_report, acct, item_no, qty, price) VALUES (?,?,?,?,?,?,?,?,?)",
                      (o, "TEST inquiry-report6.csv", i, "EM", ACCT, "AXYEG2", style, 1000, 10.0))
        a = file("ar_pdf", "TEST Reports 2.pdf")
        c.execute("INSERT INTO t_ar_detail (source_file_id, source_file, row_no, acct, dv, ship, cust_code, terms, slm) VALUES (?,?,?,?,?,?,?,?,?)",
                  (a, "TEST Reports 2.pdf", 1, ACCT, "EM", "AR2SHIP", "EM" + ACCT, "A4 NET 60 DAYS", "J35 JACK/AARON"))
        r = file("reasons", "TEST reasons.csv")
        for i, (code, desc) in enumerate(reasons.items(), 1):
            c.execute("INSERT INTO t_reason_list (source_file_id, source_file, row_no, code, description) VALUES (?,?,?,?,?)", (r, "TEST reasons.csv", i, code, desc))
    yield


def build(tmp_path, mutate=None):
    wb = openpyxl.load_workbook(TEMPLATE)
    reasons = _template_reasons()
    pl = wb["PRICE LIST"]
    for i, (_, style, *_r) in enumerate(MEMOS, 2):
        pl.cell(i, 1, style)
        pl.cell(i, 2, f"Item {style}")
    b = wb["Batch"]
    b["D1"], b["D2"], b["D4"] = "10-90001", DAY, "EM"
    first = wb["1"]
    for n, (db, style, pcs, price, claim) in enumerate(MEMOS, 1):
        ws = first if n == 1 else wb.copy_worksheet(first)
        ws.title = str(n)
        ws["B4"], ws["B6"], ws["G4"], ws["G6"], ws["H2"] = NAME, ACCT, NAME, SHIP, DAY
        ws["A9"], ws["C9"], ws["E9"], ws["F9"] = reasons[REASON_CODE], REASON_CODE, "JACK/AARON", "NET 60 DAYS"
        reason = reasons[REASON_CODE]
        ws["E10"] = reason + claim
        ws["A13"], ws["D13"], ws["G13"] = style, pcs, price
        ws["A40"] = db
        r = 7 + n
        b.cell(r, 2, NAME)
        b.cell(r, 3, DAY)
        b.cell(r, 4, pcs * price)
        b.cell(r, 5, db)
        b.cell(r, 6, claim.replace(" : ", ""))
    if mutate:
        mutate(wb)
    path = tmp_path / "batch.xlsx"
    wb.save(path)
    return path


def run(tmp_path, mutate=None):
    res = recheck.run(build(tmp_path, mutate))
    return {r["tag"]: r["count"] for r in res["results"] if r["count"]}, res


def test_clean_batch_has_no_errors(tmp_path):
    tags, res = run(tmp_path)
    assert tags == {}, [(i["where"], i["msg"]) for r in res["results"] for i in r["issues"]]
    assert res["stopped"] is None and res["skipped"] == 0


def test_line_formula_in_either_order_is_accepted(tmp_path):
    tags, _ = run(tmp_path, lambda wb: wb["1"].__setitem__("I13", "=D13*G13"))
    assert tags == {}


def test_no_sale_report_uploaded_gives_a_note_not_errors(tmp_path, monkeypatch):
    real = recheck._ref_data
    def no_sale_report():
        ref = dict(real())
        ref.update(sale_accts=set(), order_files=[], by_report={}, orders={}, sale_ship={})
        return ref
    monkeypatch.setattr(recheck, "_ref_data", no_sale_report)
    tags, res = run(tmp_path, lambda wb: (wb["1"].__setitem__("A13", "NOPE"), wb["1"].__setitem__("G6", "AZZZ9")))
    assert "P4" not in tags and "P10" not in tags and "SHIPCODE" not in tags
    assert any(n["title"] == "Sale Report not uploaded" for n in res["notes"])


def test_duplicate_item_with_a_different_price_is_not_an_error(tmp_path):
    tags, _ = run(tmp_path, lambda wb: (wb["1"].__setitem__("A14", "TST100"), wb["1"].__setitem__("G14", 9.0)))
    assert "P9" not in tags


def test_autofill_gives_the_memo_details_of_a_db_number(tmp_path):
    from app import autofill
    got = autofill.lookup(["DB#900001", "IN#NOPE"], "EM")
    assert got["DB#900001"] == {"found": True, "acct": ACCT, "name": NAME, "div": "EM", "claim": "SPA-1", "total": 100.0,
                                "salesman": "JACK/AARON", "terms": "NET 60 DAYS", "shipAcct": SHIP}
    assert got["IN#NOPE"] == {"found": False}


def test_every_protocol_runs_and_is_ordered(tmp_path):
    _, res = run(tmp_path)
    assert [r["n"] for r in res["results"]] == [n for n, *_ in recheck.CHECKS]
    assert all("stage" in r for r in res["results"])


def test_batch_number_empty_passes_and_written_is_only_a_note(tmp_path):
    tags, res = run(tmp_path, lambda wb: wb["Batch"].__setitem__("D1", None))
    assert tags == {} and not any(n["where"] == "Batch!D1" for n in res["notes"])
    tags, res = run(tmp_path, lambda wb: wb["Batch"].__setitem__("D1", "abc"))
    assert tags == {} and any(n["where"] == "Batch!D1" for n in res["notes"])


def test_same_style_different_price_is_only_a_note(tmp_path):
    def same_style(wb):
        wb["2"]["A13"] = wb["1"]["A13"].value
        wb["2"]["G13"] = wb["1"]["G13"].value + 0.5
    base, _ = run(tmp_path)
    _, res = run(tmp_path, same_style)
    assert any(n["title"] == "Same style, different price" for n in res["notes"])
    _, clean = run(tmp_path)
    assert not any(n["title"] == "Same style, different price" for n in clean["notes"])


def test_spaces_in_reason_sheet_is_only_a_note(tmp_path):
    tags, res = run(tmp_path, lambda wb: wb["REASON"].__setitem__("B3", wb["REASON"]["B3"].value + " "))
    # the template's REASON sheet already has a few such cells; a new one is one more note, never an error
    assert tags == {} and any(n["where"] == "REASON!B3" for n in res["notes"])


def test_column_width_difference_is_only_a_note(tmp_path):
    tags, res = run(tmp_path, lambda wb: (wb["1"].column_dimensions["E"].__setattr__("width", 20), wb["Batch"].column_dimensions["B"].__setattr__("width", 50)))
    assert tags == {} and any(n["where"] == "Memo 1!E1" for n in res["notes"]) and any(n["where"] == "Batch!B1" for n in res["notes"])


def test_claim_number_may_have_spaces_and_symbols(tmp_path):
    def m(wb):
        wb["1"]["E10"] = wb["1"]["E10"].value.split(" : ")[0] + " : CR 510706/A#1"
        wb["Batch"]["F8"] = "CR 510706/A#1"
    tags, _ = run(tmp_path, m)
    assert "CLAIM" not in tags and "MATH" not in tags


BREAKS = {  # name: (mutation, protocol tag that must report it)
    "format label case": (lambda wb: wb["Batch"].__setitem__("A1", "v3"), "FORMAT"),
    "memo label case": (lambda wb: wb["1"].__setitem__("A8", "Reason"), "FORMAT"),
    "memo header row height changed": (lambda wb: wb["1"].row_dimensions[4].__setattr__("height", 120), "FORMAT"),
    "memo item line height changed": (lambda wb: wb["1"].row_dimensions[13].__setattr__("height", 80), "FORMAT"),
    "batch row height changed": (lambda wb: wb["Batch"].row_dimensions[9].__setattr__("height", 40), "FORMAT"),
    "extra sheet": (lambda wb: wb.create_sheet("Extra"), "FORMAT"),
    "division small letter": (lambda wb: wb["Batch"].__setitem__("D4", "Ev"), "P19a"),
    "division not in ageing": (lambda wb: wb["Batch"].__setitem__("D4", "EV"), "P19"),
    "division empty": (lambda wb: wb["Batch"].__setitem__("D4", None), "P19a"),
    "batch date empty": (lambda wb: wb["Batch"].__setitem__("D2", None), "REQUIRED"),
    "company empty": (lambda wb: wb["Batch"].__setitem__("B8", None), "REQUIRED"),
    "no hashtag": (lambda wb: wb["Batch"].__setitem__("E8", "DB 900001"), "P2"),
    "db not in ageing": (lambda wb: (wb["Batch"].__setitem__("E8", "DB#999999"), wb["1"].__setitem__("A40", "DB#999999")), "AGEING"),
    "batch amount differs from ageing": (lambda wb: wb["Batch"].__setitem__("D8", 99.0), "AGEING"),
    "batch company differs from ageing": (lambda wb: wb["Batch"].__setitem__("B8", "OTHER CO"), "AGEING"),
    "batch claim differs from ageing": (lambda wb: wb["Batch"].__setitem__("F8", "SPA-9"), "AGEING"),
    "db under another account": (lambda wb: wb["1"].__setitem__("B6", "AM0000"), "AGEING"),
    "memo date differs": (lambda wb: wb["2"].__setitem__("H2", datetime.datetime(2026, 1, 1)), "DATE"),
    "batch date differs": (lambda wb: wb["Batch"].__setitem__("C8", datetime.datetime(2026, 1, 1)), "DATE"),
    "memo sold to differs": (lambda wb: wb["2"].__setitem__("B4", "OTHER NAME"), "MEMOCOMPANY"),
    "batch company not the ageing name": (lambda wb: wb["Batch"].__setitem__("B9", "OTHER CO"), "AGEING"),
    "memo sold to differs from batch company": (lambda wb: wb["2"].__setitem__("B4", "OTHER NAME"), "MEMOCOMPANY"),
    "ship code not in sale report nor report 2": (lambda wb: [wb[n].__setitem__("G6", "AZZZ9") for n in ("1", "2")], "SHIPCODE"),
    "salesman differs from report 2": (lambda wb: wb["1"].__setitem__("E9", "OTHER PERSON"), "CUSTCODE"),
    "terms differ from report 2": (lambda wb: wb["1"].__setitem__("F9", "NET 10 DAYS"), "P5"),
    "cust code not in report 2": (lambda wb: wb["2"].__setitem__("B6", "AM0001"), "CUSTCODE"),
    "account in small letters": (lambda wb: wb["1"].__setitem__("B6", "am9286"), "CAPS"),
    "batch row number deleted": (lambda wb: wb["Batch"].__setitem__("A9", None), "SEQ"),
    "memo sheet gap": (lambda wb: wb.__setitem__("3", wb.copy_worksheet(wb["1"])) if False else setattr(wb["2"], "title", "3"), "SEQ"),
    "claim double spacing around colon": (lambda wb: wb["1"].__setitem__("E10", wb["1"]["E10"].value.replace(" : ", "  :  ")), "CLAIM"),
    "claim wrong": (lambda wb: wb["1"].__setitem__("E10", "something else"), "CLAIM"),
    "ship via filled": (lambda wb: wb["1"].__setitem__("H9", "UPS"), "EMPTY"),
    "return filled": (lambda wb: wb["Batch"].__setitem__("D3", "x"), "EMPTY"),
    "work sheet filled": (lambda wb: wb["Work2"].__setitem__("A1", "x"), "WORK"),
    "extra formula": (lambda wb: wb["1"].__setitem__("K13", "=5"), "P6"),
    "template formula removed": (lambda wb: wb["1"].__setitem__("I13", 50), "P7"),
    "line formula with + instead of *": (lambda wb: wb["1"].__setitem__("I13", "=D13+G13"), "P7"),
    "space in a cell": (lambda wb: wb["1"].__setitem__("B4", " " + NAME), "P1"),
    "item not in sale report": (lambda wb: wb["1"].__setitem__("A13", "NOPE"), "P4"),
    "qty above sale report": (lambda wb: wb["1"].__setitem__("D13", 5000), "P10"),
    "price above sale report": (lambda wb: wb["1"].__setitem__("G13", 99.0), "P10"),
    "negative pcs": (lambda wb: wb["1"].__setitem__("D13", -1), "POSITIVE"),
    "repeated db, sum not equal to ageing": (lambda wb: (wb["Batch"].__setitem__("E9", "DB#900001"), wb["2"].__setitem__("A40", "DB#900001")), "AGEING"),
    "empty line between items": (lambda wb: (wb["1"].__setitem__("A13", None), wb["1"].__setitem__("A14", "TST100")), "GAPS"),
    "price list description empty": (lambda wb: wb["PRICE LIST"].__setitem__("B2", None), "PRICELIST"),
    "extra item on price list": (lambda wb: (wb["PRICE LIST"].__setitem__("A60", "EXTRA1"), wb["PRICE LIST"].__setitem__("B60", "Extra")), "PLMEMOS"),
    "memo item missing on price list": (lambda wb: [wb["PRICE LIST"].__setitem__(c, None) for c in ("A2", "B2")], "PLMEMOS"),
    "duplicate item on memo with the same price": (lambda wb: (wb["1"].__setitem__("A14", "TST100"), wb["1"].__setitem__("G14", wb["1"]["G13"].value)), "P9"),
    "salesman empty": (lambda wb: wb["1"].__setitem__("E9", None), "P5"),
    "blocked reason code": (lambda wb: (wb["1"].__setitem__("C9", "SP"), wb["1"].__setitem__("A9", "Short Payment")), "BLOCKED"),
    "blocked reason text only": (lambda wb: wb["1"].__setitem__("A9", "Bankruptcy"), "BLOCKED"),
    "reason differs":(lambda wb: wb["1"].__setitem__("A9", "Own words"), "P5"),
    "account name not ageing": (lambda wb: [wb[n].__setitem__(c, "WRONG NAME") for n in ("1", "2") for c in ("B4", "G4")], "P8"),
}


@pytest.mark.parametrize("name", sorted(BREAKS))
def test_protocol_catches(tmp_path, name):
    mutate, tag = BREAKS[name]
    tags, res = run(tmp_path, mutate)
    assert tag in tags, f"{name}: expected {tag}, got {tags}"



def test_dnt_claim_may_start_with_a_percentage():
    from app.recheck import _claim_no
    assert _claim_no("2% Damage Allowance : DNT1655357", "Damage Allowance") == "DNT1655357"
    assert _claim_no("6 % Co-op Allowance : DNT1630970", "Co-op Allowance") == "DNT1630970"
    assert _claim_no("2% Damage Allowance : ABC1", "Damage Allowance") is None
    assert _claim_no("2% Damage Allowance", "Damage Allowance") is None


def test_slm_match_skips_first_word_of_report():
    from app.recheck import _slm_match
    assert _slm_match("ISAAC ASH", "AAA ISAAC ASH")
    assert _slm_match("JACK", "J35 JACK/AARON")
    assert not _slm_match("ASH", "AAA ISAAC ASH")
    assert not _slm_match("OTHER PERSON", "AAA ISAAC ASH")


def _same_db(wb):
    wb["Batch"]["E9"] = "DB#900001"
    wb["2"]["A40"] = "DB#900001"


def test_repeated_db_is_grouped_and_summed_to_the_ageing(tmp_path, monkeypatch):
    """DB#900001 on two Batch rows / memos (100 + 50): the sum must equal the Ageing total (150)."""
    real = recheck._ref_data
    def ageing_150():
        ref = dict(real())
        ref["ageing"] = {**ref["ageing"], "DB#900001": [{**ref["ageing"]["DB#900001"][0], "total": 150.0, "claim_no": "SPA-1"}]}
        return ref
    monkeypatch.setattr(recheck, "_ref_data", ageing_150)
    tags, res = run(tmp_path, _same_db)
    msgs = [i["msg"] for r in res["results"] for i in r["issues"]]
    assert "P3" not in tags and "DUPLICATE" not in tags, msgs
    assert not any("Ageing Total" in m or "add up" in m for m in msgs), msgs   # only the claim # of row 2 differs (test data)
