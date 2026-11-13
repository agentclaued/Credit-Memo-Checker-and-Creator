"""The Create-batch export is a copy of the template: only template formulas, nothing the recheck calls a format or formula error."""
import openpyxl
import pytest
from fastapi.testclient import TestClient

from app import recheck, template_rules as T
from app.database import init_schema
from app.main import app


@pytest.fixture(autouse=True)
def schema():
    init_schema()          # the scratch database (tests/conftest.py) starts empty

DRAFT = {"batch": {"no": "10-90001", "date": "2026-10-16", "return": "", "division": "EM"},
         "memos": [{"date": "2026-10-16", "soldTo": "TEST CO", "soldAcct": "AM9286", "shipTo": "TEST CO", "shipAcct": "AXYEG2", "code": "DA",
                    "salesman": "JACK/AARON", "terms": "NET 60 DAYS", "via": "", "claim": "SPA-1", "db": "DB#900001", "gst": "", "freight": "", "adj": "",
                    "lines": [{"sku": "TST100", "qty": "10", "price": "10"}] + [{"sku": "", "qty": "", "price": ""}] * 20}]}


def test_export_is_the_template_with_only_template_formulas(tmp_path):
    c = TestClient(app)
    res = c.post("/api/export", json=DRAFT)
    assert res.status_code == 200 and "spreadsheetml" in res.headers["content-type"]
    path = tmp_path / "out.xlsx"
    path.write_bytes(res.content)
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == T.SHEETS
    assert wb["Batch"]["D8"].value == 100 and wb["Batch"]["D28"].value == "=SUM(D8:D27)"
    assert wb["1"]["I13"].value == "=G13*D13" and wb["1"]["D13"].value == 10
    tags = {r["tag"]: r["count"] for r in recheck.run(path)["results"] if r["count"]}
    assert not {"P6", "P7", "FORMAT"} & set(tags), tags


def test_export_without_memos_is_refused():
    res = TestClient(app).post("/api/export", json={"batch": {}, "memos": []})
    assert res.status_code == 400


def test_export_uses_the_batch_rows_typed_or_pasted_on_screen(tmp_path):
    rows = [{"company": "MENARD INC YARD# 3039", "date": "9/29/2026", "amount": " $20,000.00 ", "db": "DB#534583", "claim": "PRNT0650875", "rt": "", "rtAmt": "", "disc": "", "over": ""},
            {"company": "MENARD INC YARD# 3039", "date": "9/29/2026", "amount": "$60.00", "db": "IN#5345368", "claim": "", "rt": "", "rtAmt": "", "disc": "", "over": ""}]
    rows += [{k: "" for k in rows[0]}] * 18
    res = TestClient(app).post("/api/export", json={**DRAFT, "batchRows": rows})
    assert res.status_code == 200
    path = tmp_path / "out.xlsx"
    path.write_bytes(res.content)
    ws = openpyxl.load_workbook(path)["Batch"]
    assert ws["D8"].value == 20000 and ws["D9"].value == 60 and ws["E9"].value == "IN#5345368" and ws["F8"].value == "PRNT0650875"
    assert ws["C8"].value.strftime("%m/%d/%Y") == "09/29/2026" and ws["D10"].value is None and ws["D28"].value == "=SUM(D8:D27)"


def test_export_uses_the_price_list_typed_or_pasted_on_screen(tmp_path):
    res = TestClient(app).post("/api/export", json={**DRAFT, "priceList": [{"item": "6160", "desc": "BIONIC LIGHT"}, {"item": "", "desc": ""}, {"item": "1258ENT", "desc": "100' XHOSE"}]})
    assert res.status_code == 200
    path = tmp_path / "out.xlsx"
    path.write_bytes(res.content)
    ws = openpyxl.load_workbook(path)["PRICE LIST"]
    assert [ws["A2"].value, ws["B2"].value, ws["A3"].value] == ["6160", "BIONIC LIGHT", "1258ENT"]


def _draft(n):
    memo = DRAFT["memos"][0]
    return {**DRAFT, "memos": [{**memo, "db": f"DB#9{i:05d}", "claim": f"C{i}"} for i in range(n)],
            "batchRows": [{"company": "TEST CO", "date": "2026-10-16", "amount": "100", "db": f"DB#9{i:05d}", "claim": f"C{i}"} for i in range(n)]}


def test_more_than_20_memos_grow_the_batch_and_move_the_total(tmp_path):
    res = TestClient(app).post("/api/export", json=_draft(22))
    assert res.status_code == 200
    path = tmp_path / "out.xlsx"
    path.write_bytes(res.content)
    wb = openpyxl.load_workbook(path)
    b = wb["Batch"]
    assert [b["A28"].value, b["A29"].value, b["D29"].value] == [21, 22, 100]
    assert b["D30"].value == "=SUM(D8:D29)" and b["A30"].value is None
    assert len([n for n in wb.sheetnames if n.isdigit()]) == 22
    tags = {r["tag"]: r["count"] for r in recheck.run(path)["results"] if r["count"]}
    assert not {"SEQ", "FORMAT", "P6"} & set(tags), tags       # no 20-memo limit, the moved total is still the template formula


def test_memos_keep_the_zoom_of_memo_1(tmp_path):
    res = TestClient(app).post("/api/export", json=_draft(3))
    path = tmp_path / "out.xlsx"
    path.write_bytes(res.content)
    wb = openpyxl.load_workbook(path)
    zooms = {wb[n].sheet_view.zoomScale for n in ("1", "2", "3")}
    assert len(zooms) == 1


def test_price_list_stays_empty_when_nothing_is_sent(tmp_path):
    res = TestClient(app).post("/api/export", json=DRAFT)
    path = tmp_path / "out.xlsx"
    path.write_bytes(res.content)
    ws = openpyxl.load_workbook(path)["PRICE LIST"]
    assert all(c.value is None for r in ws.iter_rows(min_row=2) for c in r)


def test_autofill_reports_unknown_db_as_not_found():
    res = TestClient(app).post("/api/autofill", json={"refs": ["DB#1", "IN#2"], "division": "EM"})
    assert res.status_code == 200 and res.json() == {"DB#1": {"found": False}, "IN#2": {"found": False}}


def test_create_page_has_the_batch_controls():
    html = TestClient(app).get("/memo").text
    for needle in ('id="dbOnly" checked', 'id="reasonAll"', 'id="clearBatch"', 'id="clearPrice"', 'id="exportBtn"'):
        assert needle in html, needle
