"""Recheck of a filled Credit Memo Batch against the 10 recheck points (credit-memo-protocols.pdf).

The uploaded workbook is read directly with openpyxl (formulas and raw cell text are needed: spaces, formulas).
Ageing and Sale Report data come from the rows already imported into SQLite. Nothing is written back.
"""
import re
from collections import defaultdict
from datetime import date, datetime
from functools import lru_cache

import openpyxl
from openpyxl.utils import get_column_letter

from .blocked_reasons import is_blocked
from .database import DB_PATH, ROOT, transaction
from .template_rules import TEMPLATE_FILE

TEMPLATE_PATH = ROOT / "templates_xlsx" / TEMPLATE_FILE

CHECKS = [
    (1, "No spaces in cells", "No space before or after the text in any cell, on every sheet. A cell with only a space is also wrong."),
    (2, "DB# with hashtag", "DB#123456 (or IN#123456). The # is a must. Same number on the Batch sheet and the memo sheet."),
    (3, "DB amount = Ageing", "The DB number must exist in the Ageing and the memo amount must equal the Ageing amount. A DB# on several memos: the memo totals are added together and must equal the Ageing amount."),
    (4, "Item in sale report", "Every item on the memo must be in the sale report of the same account."),
    (5, "Salesman, Terms, Reason", "All four filled. Salesman matches the account: the names are split on / & , ; and at least one name must be the same (JACK/AARON matches AARON; case and extra spaces ignored). Terms: the Cust Code is built from the Batch DIVISION (D4) + SOLD TO account (B6) and the Terms must match its Terms in Report 2. Reason and code are exactly (case sensitive) as in the REASON sheet; a code that is not there is looked up in the Reason codes list of the DataBase Room."),
    (6, "No extra formula", "Only the Batch total cell has a formula on the Batch sheet. Nothing else may be added anywhere."),
    (7, "Template formulas kept", "VLOOKUP, price x quantity (=G13*D13 or =D13*G13, both accepted), SUBTOTAL and TOTAL must stay on every memo. A typed number instead of the formula shows \"No formula =G13*D13 found. Manual value 120 was typed\"."),
    (8, "Account name = Ageing", "SOLD TO name and account, and the Batch COMPANY, must be the same as in the Ageing."),
    (9, "No duplicate item", "One item once on the PRICE LIST. On a memo the same item twice is an error only when the price is the same too."),
    (10, "Qty and price within sale report", "Memo quantity and price must not be more than the sale report. Less is allowed."),
    (11, "REASON sheet = database", "The REASON sheet of the file must match the Reason codes list imported in the DataBase Room (case does not matter). Then every memo reason and code must match that REASON sheet exactly (case sensitive)."),
    (12, "Format unchanged", "The template sheets, labels and layout are intact: required sheets, no extra sheets, Batch and memo labels, Batch total formula. A column width that differs from the template is only a note."),
    (13, "Memo sheets in sequence", "At least 1 memo sheet (no upper limit), numbered 1, 2, 3 without gaps. Batch rows numbered 1, 2, 3 and as many as memo sheets."),
    (14, "Claim # format", "CUSTOMER CLAIM # is '<Reason> : <claim no.>', or just '<Reason>' when there is no claim number. The claim no. is the Batch CLAIM # of that memo (any text, numbers, symbols); one space each side of the colon."),
    (15, "Ship via / Return empty", "The RETURN cell on the Batch sheet and SHIPPED VIA on every memo must be empty."),
    (16, "Work sheets empty", "Work1, Work2 and Work3 must have no data."),
    (17, "Same date", "Every Batch row date and every memo date is filled and equals the batch DATE."),
    (18, "Math and cross-checks", "Line amount = PCS x price, SUBTOTAL, TOTAL, Batch total, and Batch row = memo (amount, claim #)."),
    (19, "Division in capitals", "The DIVISION on the Batch sheet is two letters, both in capitals (EV, not Ev)."),
    (20, "Required fields filled", "Batch number and DATE, COMPANY on every Batch row, and SOLD TO name and account on every memo are filled in."),
    (21, "Account in capitals", "SOLD TO and SHIP TO account numbers are written in capitals with no spaces (AM9286). Style numbers in small letters show as a note."),
    (22, "Positive numbers", "PCS is a whole number above zero, price and amounts are above zero, and GST/HST, freight and adjustment are numbers."),
    (23, "No duplicate code", "A reason code may be on the REASON sheet only once. (A DB# may repeat on the Batch: see Batch = Ageing.)"),
    (24, "No empty line between items", "On a memo the items follow each other with no empty line between them."),
    (25, "PRICE LIST complete", "Every item on the PRICE LIST has a description."),
    (26, "Batch = Ageing", "Every Batch row: DB# is in the Ageing Ref#; AMOUNT = Ageing Total (a DB# on several rows: group by DB#, the sum of the AMOUNTs = Ageing Total); COMPANY = Ageing Name; CLAIM # = Ageing Claim#. (The DATE must be the batch DATE: see Same date.)"),
    (29, "Memo company = Batch company = Ageing", "SOLD TO name (B4) of every memo is the same as the COMPANY of its own Batch row (memo 3 = Batch row 3) and as the Ageing Name of the account (case does not matter)."),
    (30, "Ship code in Sale Report / Report 2", "SHIP TO account (G6) on every memo must be an Acct ShipCode of the Sale Report, or else a Ship Code of Report 2 (A/R PDF), for the account Batch DIVISION (D4) + SOLD TO account (B6)."),
    (32, "Division = Ageing Div", "The Batch DIVISION (D4) must be a Div of the Ageing for the SOLD TO account of the memos (the Ageing is grouped by Acct and Div)."),
    (31, "Salesman = Report 2", "Cust Code = Batch DIVISION (D4) + SOLD TO account (B6) (EM + AM9286 = EMAM9286) must exist in Report 2; the memo SALESMAN (E9) must match its SLM by name, at least one name the same, a leading code such as J35 ignored (J35 JACK/AARON matches JACK/AARON or AARON) (the Terms are checked by Salesman, Terms, Reason)."),
    (34, "Reason not allowed", "No credit memo may use these reason codes: Short Payment (SP), Allow Deductions (AL), Short Shipped (SS), Bankruptcy (BK), Partial Payment (PT), Cancel (CX), Unauthorized Deduction (UD). Checked on the reason and the code of every memo."),
    (33, "PRICE LIST = memos", "The PRICE LIST has exactly the items of the memo sheets of the batch: every memo item is on the PRICE LIST and no other item is."),
]
# The checks are ordered in stages. With STOP_AT_FAILED_STAGE a later stage is hidden while an earlier one has errors; off = every error is shown.
NAMES = {c[0]: c[1] for c in CHECKS}
STOP_AT_FAILED_STAGE = False
STAGES = [("Batch format", [12]), ("Batch values and Ageing", [2, 26, 17, 29, 30, 31, 32])]   # all other checks form the last stage
TAGS = {11: "D_P11", 12: "FORMAT", 13: "SEQ", 14: "CLAIM", 15: "EMPTY", 16: "WORK", 17: "DATE", 18: "MATH", 19: "P19a",
        20: "REQUIRED", 21: "CAPS", 22: "POSITIVE", 23: "DUPLICATE", 24: "GAPS", 25: "PRICELIST", 26: "AGEING", 29: "MEMOCOMPANY", 30: "SHIPCODE", 31: "CUSTCODE", 32: "P19", 33: "PLMEMOS", 34: "BLOCKED"}   # database protocol and the PROTOCOLS.md rules; the others show as P1..P10
FIXES = {
    12: "Use the template as it is; do not rename, delete or add sheets or labels",
    13: "Number the memo sheets 1, 2, 3... and keep one Batch row per memo",
    14: "Write it like: Reason : claim number (or only the Reason)",
    15: "Clear the cell",
    16: "Clear the Work sheets",
    17: "Use the batch date in every date cell",
    18: "Correct the amount so the math agrees",
    32: "Use the Div of the Ageing for this account in the Batch DIVISION",
    19: "Type the division, both letters in capitals",
    20: "Fill it in",
    21: "Write the account in capitals without spaces",
    22: "Use a whole number above zero for PCS and an amount above zero",
    23: "Keep each DB# and each reason code only once",
    24: "Move the items up so there is no empty line between them",
    25: "Add the description of the item on the PRICE LIST",
    34: "Choose another reason code; this one cannot be used for a credit memo",
    33: "Add the missing memo items to the PRICE LIST and delete the items that are on no memo",
    26: "Copy the value from the Ageing",
    31: "Copy the salesman and terms of the Cust Code from Report 2",
    30: "Use a ship code that is on the Sale Report or Report 2 for this account",
    29: "Copy the Ageing name into SOLD TO on every memo and into the Batch COMPANY",
    1: "Remove the extra space (press Delete)",
    2: "Write it with the hashtag, like DB#123456, the same on Batch and memo",
    3: "Check the DB in the Ageing first; the memo amount must be the same",
    4: "Use an item that was sold to this account",
    5: "Fill it in; use the code and exact text from the REASON sheet",
    6: "Delete the extra formula and type the value by hand",
    7: "Restore the template formula",
    8: "Copy the account name and number from the Ageing",
    9: "Keep each item only once",
    10: "Use a quantity and price not more than the sale report",
    11: "Make the REASON sheet the same as the Reason codes database",
}
DB_RE = re.compile(r"^(DB|IN)#\d+$")
REQUIRED_SHEETS = ["REASON", "PRICE LIST", "Work1", "Work2", "Work3", "Batch"]
MONEY_TOL = 0.005


@lru_cache(maxsize=1)
def _golden():
    """The template workbook (templates_xlsx): labels, formulas and formatting are compared with it, case sensitive."""
    return openpyxl.load_workbook(TEMPLATE_PATH)


def _fmt(c):
    fill = c.fill.fgColor if c.fill and c.fill.fill_type else None
    color = None if fill is None else (fill.rgb if isinstance(fill.rgb, str) else f"theme{fill.theme}")
    return (c.number_format, bool(c.font.b), color)


DEFAULT_WIDTH, DEFAULT_HEIGHT = 8.43, 15.0
WIDTH_TOL, HEIGHT_TOL = 0.5, 1.0     # Excel rounds sizes a little on different screens; a visible change is larger


def _widths(ws):
    """Column index -> width, with Excel's grouped columns (one entry for A:C) expanded."""
    out = {}
    for d in ws.column_dimensions.values():
        if d.width is not None and d.min:
            for idx in range(d.min, (d.max or d.min) + 1):
                out[idx] = d.width
    return out


def _height(ws, r):
    h = ws.row_dimensions[r].height if r in ws.row_dimensions else None
    return DEFAULT_HEIGHT if h is None else h


def _date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = _txt(v).strip()
    for f in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def _claim_no(claim: str, reason: str):
    """'<Reason> : <claim no.>' -> claim no.; just '<Reason>' -> ''; anything else -> None (wrong format)."""
    if claim == reason:
        return ""
    m = re.fullmatch(re.escape(reason) + r" : (\S.*)", claim)       # one space each side of the colon; the claim no. is free text (letters, numbers, symbols, spaces)
    if m:
        return m.group(1)
    m = re.fullmatch(r"\d+(?:\.\d+)? ?% " + re.escape(reason) + r" : (\S.*)", claim)   # a DNT claim may start with a percentage: "2% Damage Allowance : DNT1655357"
    return m.group(1) if m and "DNT" in m.group(1).upper() else None


_WS = re.compile(r"\s+")


def _line_formula(v, r) -> bool:
    """Line total of a memo row: price x quantity, either order (=G13*D13 or =D13*G13)."""
    return v in (f"=G{r}*D{r}", f"=D{r}*G{r}")


def _slm_names(v, skip_first=False) -> set:
    """Salesman names of a cell: split on / & , ; and drop a leading code such as J35 (J35 JACK/AARON -> JACK, AARON).
    skip_first (report side): always drop the first word, keeping the rest (AAA ISAAC ASH -> ISAAC ASH)."""
    out = set()
    for part in re.split(r"[/&,;]", _norm(v)):
        part = part.strip()
        if skip_first and " " in part:
            part = part.split(" ", 1)[1].strip()
        else:
            part = re.sub(r"^\S*\d\S*\s+", "", part)
        if part:
            out.add(part)
    return out


def _slm_match(memo, report) -> bool:
    """True when the memo salesman equals a report salesman once the report's first word (its code) is skipped."""
    return bool(_slm_names(memo) & (_slm_names(report, True) | _slm_names(report)))


def _norm(v) -> str:
    return _WS.sub(" ", str(v or "")).strip().upper()


def _txt(v) -> str:
    return "" if v is None else str(v)


def _num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


_REF_CACHE = {}


def _ref_data():
    """Latest ageing file + all sale report rows from SQLite. Reloaded only when the database file changed (an import or edit)."""
    st = DB_PATH.stat()
    key = (st.st_mtime_ns, st.st_size)
    if _REF_CACHE.get("key") != key:
        _REF_CACHE.update(key=key, data=_load_ref_data())
    return _REF_CACHE["data"]


def _load_ref_data():
    """Latest ageing file + all sale report rows from SQLite."""
    with transaction() as conn:
        f = conn.execute("SELECT id, filename FROM source_files WHERE kind='ageing' AND status='imported' ORDER BY id DESC LIMIT 1").fetchone()
        ageing, analyst, ageing_div = {}, {}, defaultdict(set)
        names = {}
        if f:
            for r in conn.execute("SELECT ref_no, acct, name, total, claim_no, date FROM t_ageing_full WHERE source_file_id=? AND ref_no IS NOT NULL", (f["id"],)):
                ageing.setdefault(_norm(r["ref_no"]), []).append(dict(r))
                names.setdefault(_norm(r["acct"]), r["name"])
            for r in conn.execute("SELECT acct, div FROM t_ageing_full WHERE source_file_id=? AND acct IS NOT NULL AND div IS NOT NULL GROUP BY acct, div", (f["id"],)):
                ageing_div[_norm(r["acct"])].add(_norm(r["div"]))
            for r in conn.execute("SELECT acct, company, terms, slm FROM t_account_analyst WHERE source_file_id=?", (f["id"],)):
                analyst[_norm(r["acct"])] = dict(r)
        orders = defaultdict(list)    # (source_file, item) -> [(acct, qty, price)]
        by_report = defaultdict(list)  # (SaleReport account e.g. AM9286, item) -> [(account, qty, price)]
        for r in conn.execute("SELECT source_file, sale_report, acct, item_no, qty, price FROM t_orders"):
            line = (_norm(r["acct"]), r["qty"] or 0, r["price"])
            if r["sale_report"]:
                by_report[(_norm(r["sale_report"]), _norm(r["item_no"]))].append(line)
            else:                  # older import without the SaleReport column
                orders[(r["source_file"], _norm(r["item_no"]))].append(line)
        sale_ship = defaultdict(set)   # Dv+SaleReport (EMAM9286) -> {Acct ShipCode}
        for r in conn.execute("SELECT COALESCE(dv,'') || COALESCE(sale_report,'') AS k, acct FROM t_orders WHERE acct IS NOT NULL GROUP BY 1, 2"):
            sale_ship[_norm(r["k"])].add(_norm(r["acct"]))
        ar_ship = defaultdict(set)     # Dv+Acct of Report 2 (A/R detail) -> {Ship Code}
        for r in conn.execute("SELECT COALESCE(dv,'') || COALESCE(acct,'') AS k, ship FROM t_ar_detail WHERE ship IS NOT NULL GROUP BY 1, 2"):
            ar_ship[_norm(r["k"])].add(_norm(r["ship"]))
        sale_accts = {_norm(r[0]) for r in conn.execute("SELECT DISTINCT sale_report FROM t_orders WHERE sale_report IS NOT NULL")}
        sale_accts |= {_norm(r[0]) for r in conn.execute("SELECT DISTINCT acct FROM t_orders WHERE sale_report IS NULL AND acct IS NOT NULL")}   # older import without the SaleReport column
        sale_sources = defaultdict(list)   # SaleReport account -> source files it came from
        for r in conn.execute("SELECT DISTINCT sale_report, source_file FROM t_orders WHERE sale_report IS NOT NULL"):
            sale_sources[_norm(r["sale_report"])].append(r["source_file"])
        cust = defaultdict(set)        # Report 2 Cust Code (EMAM9285) -> {(Terms, SLM)}
        for q in ("SELECT code AS k, terms, slm FROM t_ar_customers WHERE code IS NOT NULL",
                  "SELECT cust_code AS k, terms, slm FROM t_ar_detail WHERE cust_code IS NOT NULL GROUP BY 1, 2, 3"):
            for r in conn.execute(q):
                cust[_norm(r["k"]).split(" ")[0][2:]].add((_norm(r["terms"]), _norm(r["slm"])))
        files = [r[0] for r in conn.execute("SELECT DISTINCT source_file FROM t_orders")]
        rf = conn.execute("SELECT id FROM source_files WHERE kind='reasons' AND status='imported' ORDER BY id DESC LIMIT 1").fetchone()
        reasons = None   # None = no Reason codes list imported
        if rf:
            reasons = {_norm(r["code"]): _txt(r["description"]) for r in conn.execute("SELECT code, description FROM t_reason_list WHERE source_file_id=? AND code IS NOT NULL", (rf["id"],))}
    return {"ageing_file": f["filename"] if f else None, "ageing": ageing, "analyst": analyst, "names": names, "ageing_div": ageing_div,
            "orders": orders, "by_report": by_report, "sale_ship": sale_ship, "cust": cust, "sale_sources": sale_sources, "sale_accts": sale_accts, "ar_ship": ar_ship, "order_files": files, "reasons": reasons}


def _ageing_find(ref, db, acct, name):
    """Ageing rows of this DB#, taken only from the account (and name) of the batch: filter Acct, then Name, then the DB#.
    Returns (rows of that account, rows of the same DB# in any account)."""
    anywhere = ref["ageing"].get(_norm(db), [])
    mine = [r for r in anywhere if _norm(r["acct"]) == acct] if acct else list(anywhere)
    if name:
        mine = [r for r in mine if _norm(r["name"]) == _norm(name)] or mine
    return mine, anywhere


def _not_found(db, acct, anywhere):
    other = sorted({_norm(r["acct"]) for r in anywhere if _norm(r["acct"]) != acct})
    return f"{db} is not in the Ageing of account {acct or '(no account)'}" + (f" (it is under {', '.join(other)})" if other else "")


def _sale_lines(ref, sold, ship, item):
    """Sale report lines of this account for an item: files named after the sold-to account, else rows of that account."""
    hit = [x for acct in dict.fromkeys((sold, ship)) if acct for x in ref["by_report"].get((acct, item), [])]
    if hit:
        return hit
    files = [f for f in ref["order_files"] if _norm(f).startswith(sold)]
    if files:
        return [x for f in files for x in ref["orders"].get((f, item), [])]
    return [x for f in ref["order_files"] for x in ref["orders"].get((f, item), []) if x[0] in (sold, ship)]


def _has_sale_report(ref, sold, ship) -> bool:
    """True when a Sale Report of this account (SOLD TO or SHIP TO) was uploaded."""
    return bool({sold, ship} & ref["sale_accts"]) or any(_norm(f).startswith(sold) for f in ref["order_files"] if sold)


def _memo_sale_files(wb, memos, ref) -> list:
    """'AM9286: file1.csv, file2.csv' for each SOLD TO account on the memos (its SaleReport value and the source files)."""
    out = []
    for acct in dict.fromkeys(_norm(wb[n]["B6"].value) for n in memos):
        if acct:
            out.append(f"{acct}: " + (", ".join(ref["sale_sources"].get(acct, [])) or "not imported"))
    return out


def run(path) -> dict:
    ref = _ref_data()
    wb = openpyxl.load_workbook(path)                    # formulas
    vals = openpyxl.load_workbook(path, data_only=True)  # cached values
    issues = {n: [] for n, *_ in CHECKS}
    notes = []            # shown as notes only; not counted as errors

    def bad(n, where, msg, fix=None, value=None, expected=None):
        sheet = where.split("!")[0]
        issues[n].append({"where": where, "msg": msg, "fix": fix or FIXES[n], "value": None if value is None else str(value), "tag": TAGS.get(n, f"P{n}"), "n": n, "title": NAMES[n], "label": TAGS.get(n, f"P{n}") if n in (19, 32) else f"P{n}",
                          "expected": None if expected is None else str(expected),
                          "sheet": sheet[5:] if sheet.startswith("Memo ") else sheet})

    def note(where, msg, fix, value, title):
        sheet = where.split("!")[0]
        notes.append({"where": where, "msg": msg, "fix": fix, "value": None if value is None else str(value), "tag": "NOTE",
                      "sheet": sheet[5:] if sheet.startswith("Memo ") else sheet, "title": title})

    # 1 ---- spaces in every cell of every sheet
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith("=") and c.value != c.value.strip():
                    msg = "Empty cell with a space in it" if not c.value.strip() else f"Leading/trailing space in \"{c.value}\""
                    fix = "Press Delete on the cell" if not c.value.strip() else "Remove the extra space"
                    if ws.title == "REASON":      # spaces here are trimmed when the codes are matched: a note, not an error
                        note(f"{ws.title}!{c.coordinate}", msg, fix, repr(c.value), "No spaces in cells")
                    else:
                        bad(1, f"{ws.title}!{c.coordinate}", msg, fix, value=repr(c.value))

    # 6 ---- formulas allowed only in the template places
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                ok = False
                if ws.title == "Batch":
                    ok = c.column_letter == "D" and re.fullmatch(r"=SUM\(D8:D\d+\)", v) is not None
                elif ws.title.strip().isdigit():
                    r = c.row
                    ok = ((c.column_letter == "E" and v == f"=VLOOKUP(A{r},'PRICE LIST'!A:B,2,0)") or
                          (c.column_letter == "I" and (_line_formula(v, r) or re.fullmatch(r"=SUM\(I13:I\d+\)", v) or
                                                       re.fullmatch(r"=SUM\(I\d+:I\d+\)", v))))
                if not ok:
                    bad(6, f"{ws.title}!{c.coordinate}", f"Extra formula {v}", "Delete the formula and type the value by hand", value=v)

    # Batch sheet rows
    batch = wb["Batch"] if "Batch" in wb.sheetnames else None
    batch_rows = {}
    if batch:
        start = next((r for r in range(1, batch.max_row + 1) if _txt(batch.cell(r, 1).value).strip() == "#"), None)
        last_key = 0
        for r in range((start or 0) + 1, batch.max_row + 1):
            n = batch.cell(r, 1).value
            if _txt(batch.cell(r, 4).value).startswith("="):
                break                         # the total row ends the Batch rows
            if not isinstance(n, (int, float)) and any(batch.cell(r, c).value not in (None, "") for c in range(2, 11)):
                n = last_key + 1              # a filled row whose number was deleted still counts (reported by the sequence check)
            if isinstance(n, (int, float)):
                last_key = int(n)
                batch_rows[int(n)] = {"row": r, "company": _txt(batch.cell(r, 2).value), "amount": _num(batch.cell(r, 4).value),
                                      "db": _txt(batch.cell(r, 5).value)}

    # REASON + PRICE LIST
    reasons = {}          # exact code -> exact description, as written in this file's REASON sheet
    reason_rows = []
    if "REASON" in wb.sheetnames:
        for i, r in enumerate(wb["REASON"].iter_rows(min_row=3, values_only=True), start=3):
            if r and r[0] not in (None, ""):
                reasons[_txt(r[0]).strip()] = _txt(r[1]).strip()
                reason_rows.append((i, _txt(r[0]).strip(), _txt(r[1]).strip()))
    # 11 (D_P11) ---- the REASON sheet of the file must match the Reason codes database (case NOT sensitive)
    sheet_ok = True
    if ref["reasons"] is None:
        bad(11, "REASON!A3", "No Reason codes list imported yet. Upload it in the DataBase Room.", value="(no list)")
        sheet_ok = False
    else:
        for i, code, desc in reason_rows:
            known = ref["reasons"].get(_norm(code))
            if known is None:
                bad(11, f"REASON!A{i}", f"Code {code} is not in the Reason codes database", value=code)
                sheet_ok = False
            elif _norm(desc) != _norm(known):
                bad(11, f"REASON!B{i}", f"Description of {code} must be \"{known}\" like the database", value=desc)
                sheet_ok = False
    codes_seen = {}       # 23 ---- a reason code only once on the REASON sheet
    for i, code, desc in reason_rows:
        if code in codes_seen:
            bad(23, f"REASON!A{i}", f"Code {code} is also on row {codes_seen[code]}", value=code)
        codes_seen.setdefault(code, i)
    if "PRICE LIST" in wb.sheetnames:
        seen = {}
        for i, r in enumerate(wb["PRICE LIST"].iter_rows(min_row=2, values_only=True), start=2):
            k = _norm(r[0]) if r else ""
            if k and not _txt(r[1] if len(r) > 1 else None).strip():          # 25 ---- description of every item
                bad(25, f"PRICE LIST!B{i}", f"Item {r[0]} has no description", value="(empty)")
            if k:
                if k in seen:
                    bad(9, f"PRICE LIST!A{i}", f"Item {r[0]} is also on row {seen[k]}", value=r[0])
                seen.setdefault(k, i)

    memos = sorted((n for n in wb.sheetnames if n.strip().isdigit()), key=int)

    # 33 (PLMEMOS) ---- the PRICE LIST holds exactly the items of the memos
    if "PRICE LIST" in wb.sheetnames and memos:
        memo_items = {}       # normalised item -> (item as written, first sheet)
        for m in memos:
            mws = wb[m]
            end = next((r for r in range(1, mws.max_row + 1) if _norm(mws.cell(r, 5).value) == "SUBTOTAL"), mws.max_row + 1)
            for r in range(13, end):
                k = _norm(mws.cell(r, 1).value)
                if k:
                    memo_items.setdefault(k, (_txt(mws.cell(r, 1).value).strip(), m))
        pl_items = {}
        for i, r in enumerate(wb["PRICE LIST"].iter_rows(min_row=2, values_only=True), start=2):
            k = _norm(r[0]) if r else ""
            if k:
                pl_items.setdefault(k, (_txt(r[0]).strip(), i))
        for k, (item, i) in pl_items.items():
            if k not in memo_items:
                bad(33, f"PRICE LIST!A{i}", f"Item {item} is on the PRICE LIST but on no memo", value=item)
        for k, (item, m) in memo_items.items():
            if k not in pl_items:
                bad(33, f"{m}!A13", f"Item {item} of memo {m} is not on the PRICE LIST", value=item)

    # 12 (FORMAT) ---- template sheets and labels
    for s in REQUIRED_SHEETS:
        if s not in wb.sheetnames:
            bad(12, s, f"Sheet {s} is missing", value="(missing)")
    for s in wb.sheetnames:
        if s not in REQUIRED_SHEETS and not s.strip().isdigit():
            bad(12, s, f"Sheet {s} is not allowed (only the template sheets and numbered memo sheets)", "Delete the sheet", value=s)
    # labels, formulas and their formatting must be exactly the template's (case sensitive); merged cells must stay
    gold = _golden()
    for gname, targets in (("Batch", ["Batch"] if batch else []), ("1", memos)):
        gws = gold[gname]
        for target in targets:
            ws = wb[target]
            where_sheet = target if gname == "Batch" else f"Memo {target}"
            # Fixed parts only: Batch rows 1-7, memo rows 1-12, and the memo footer (SUBTOTAL..DEBIT MEMO NOS.) which moves with the number of lines.
            # Unused Batch rows and unused memo lines may be removed, so rows in between are not compared. Formulas are checked by P6/P7.
            gsub = next((r for r in range(1, gws.max_row + 1) if gws.cell(r, 5).value == "SUBTOTAL"), 10 ** 6)
            usub = next((r for r in range(1, ws.max_row + 1) if ws.cell(r, 5).value == "SUBTOTAL"), None) if gname == "1" else None
            for row in gws.iter_rows():
                for gc in row:
                    if gc.value is None or str(gc.value).startswith("="):
                        continue
                    if gname == "Batch":
                        if gc.row > 7:
                            continue
                        coord = gc.coordinate
                    elif gc.row <= 12:
                        coord = gc.coordinate
                    elif gc.row >= gsub:
                        if usub is None:
                            continue
                        coord = ws.cell(usub + gc.row - gsub, gc.column).coordinate
                    else:
                        continue
                    c = ws[coord]
                    if c.value != gc.value:
                        bad(12, f"{where_sheet}!{coord}", f"Must be \"{gc.value}\" like the template (case sensitive)", value=_txt(c.value) or "(empty)", expected=gc.value)
                    elif _fmt(c) != _fmt(gc):
                        bad(12, f"{where_sheet}!{coord}", "Formatting differs from the template (number format, bold or fill colour)", "Copy the format from the template",
                            value=" / ".join(map(str, _fmt(c))))
            if gname == "1" and usub is None:
                bad(12, f"{where_sheet}!E34", "The SUBTOTAL row is missing (case sensitive)", value="(missing)")
            # column widths and row heights (header rows and the memo footer) must be the template's
            gw, uw = _widths(gws), _widths(ws)
            gmax = gws.max_column                 # once: openpyxl rescans every cell on each call
            for idx, gwid in sorted(gw.items()):
                if idx > gmax:
                    continue                          # Excel's "J:XFD" group covers the whole sheet: compare the used columns only
                uwid = uw.get(idx, DEFAULT_WIDTH)
                if abs(gwid - uwid) > WIDTH_TOL:
                    letter = get_column_letter(idx)
                    note(f"{where_sheet}!{letter}1", f"Column {letter} width is {uwid:.1f} but the template has {gwid:.1f}", "Set the column width back to the template's",
                         f"{uwid:.1f}", "Column width differs from the template")
            rows = [(r, r) for r in range(1, (8 if gname == "Batch" else 13))]
            if gname == "1" and usub is not None:
                rows += [(13, r) for r in range(13, usub)]                                   # item lines: all the height of the template's first line
                rows += [(r, usub + r - gsub) for r in range(gsub, gws.max_row + 1)]
            elif gname == "Batch":
                tot = next((r for r in range(8, ws.max_row + 1) if _txt(ws.cell(r, 4).value).startswith("=SUM")), ws.max_row + 1)
                rows += [(8, r) for r in range(8, tot)]                                      # Batch rows: the template's row height
            for gr, ur in rows:
                gh, uh = _height(gws, gr), _height(ws, ur)
                if abs(gh - uh) > HEIGHT_TOL:
                    bad(12, f"{where_sheet}!A{ur}", f"Row {ur} height is {uh:.1f} but the template has {gh:.1f}", "Set the row height back to the template's",
                        value=f"{uh:.1f}", expected=f"{gh:.1f}")
            head = {str(m) for m in gws.merged_cells.ranges if m.max_row <= (7 if gname == "Batch" else 12)}
            missing = head - {str(m) for m in ws.merged_cells.ranges}
            if missing:
                bad(12, f"{where_sheet}!{sorted(missing)[0].split(':')[0]}", f"Merged cells of the template are missing: {', '.join(sorted(missing)[:6])}", "Do not unmerge cells", value=f"{len(missing)} missing")
    if batch:
        total_cell = next((r for r in range(8, batch.max_row + 1) if re.fullmatch(r"=SUM\(D8:D\d+\)", _txt(batch.cell(r, 4).value))), None)
        if total_cell is None:
            bad(12, f"Batch!D{batch.max_row}", "The Batch total formula =SUM(D8:D…) is missing", "Restore the template formula")

    # 13 (SEQ) ---- memo sheets 1-20 in sequence, Batch rows numbered and as many as the memos
    nums = sorted(int(n) for n in memos)
    if not nums:
        bad(13, "Batch", f"There must be at least 1 memo sheet, found {len(nums)}", value=len(nums))
    elif nums != list(range(1, len(nums) + 1)):
        bad(13, "Batch", f"Memo sheets must be numbered 1, 2, 3… without gaps, found {', '.join(map(str, nums))}", value=", ".join(map(str, nums)))
    filled_rows = []          # Batch rows that hold data (COMPANY..CREDIT OVER / UNDER)
    unnumbered = []           # (row, its position) filled Batch rows whose number in column A was deleted
    if batch:
        hdr = next((r for r in range(1, batch.max_row + 1) if _txt(batch.cell(r, 1).value).strip() == "#"), None)
        if hdr:
            i = 0
            for r in range(hdr + 1, (total_cell or batch.max_row + 1)):      # rows above the total row
                has_data = any(batch.cell(r, c).value not in (None, "") for c in range(2, 11))
                if batch.cell(r, 1).value is None and not has_data:
                    continue
                i += 1
                if batch.cell(r, 1).value is None and has_data:
                    unnumbered.append((r, i))
                    bad(13, f"Batch!A{r}", f"The Batch row of memo {i} (row {r}) has no number, so memo {i} has no numbered row on the Batch sheet",
                        f"Type {i} in cell A{r}", value="(empty)")
                elif batch.cell(r, 1).value != i:
                    bad(13, f"Batch!A{r}", f"Batch row number must be {i}", value=_txt(batch.cell(r, 1).value) or "(empty)", expected=i)
                if has_data:
                    filled_rows.append(r)
            if len(filled_rows) != len(nums):
                bad(13, "Batch", f"{len(filled_rows)} Batch rows are filled but there are {len(nums)} memo sheets", value=f"{len(filled_rows)} vs {len(nums)}")

    # 19 (DIVISION) ---- Batch DIVISION: two letters, both capitals
    if batch:
        div = _txt(batch["D4"].value).strip()
        if not re.fullmatch(r"[A-Z]{2}", div):
            bad(19, "Batch!D4", "DIVISION must be two letters in capitals" if div else "DIVISION is empty", value=_txt(batch["D4"].value) or "(empty)", expected=div.upper() if div else None)

    # 32 (P19) ---- Batch DIVISION (D4) is a Div of the Ageing (grouped by Acct + Div) for the SOLD TO account of the memos
    if batch and ref["ageing_file"]:
        div = _norm(batch["D4"].value)
        accts = sorted({_norm(wb[n]["B6"].value) for n in memos if _norm(wb[n]["B6"].value)})
        for acct in accts:
            divs = ref["ageing_div"].get(acct)
            if div and divs and div not in divs:
                bad(32, "Batch!D4", f"DIVISION {div} is not an Ageing Div of account {acct} (Ageing has {', '.join(sorted(divs))})", value=div, expected=sorted(divs)[0])

    # 20 (REQUIRED) Batch number, DATE, COMPANY;  23 (DUPLICATE) DB# only once
    if batch:
        bno = _txt(batch["D1"].value).strip()
        if bno and not re.fullmatch(r"\d+-\d+", bno):
            note("Batch!D1", f"The batch number \"{bno}\" is not like 10-26973", None, bno, "Batch number format")
        if _date(batch["D2"].value) is None:
            bad(20, "Batch!D2", "The batch DATE is empty or not a date", value=_txt(batch["D2"].value) or "(empty)")
        db_rows = {}
        for r in filled_rows:
            if not _txt(batch.cell(r, 2).value).strip():
                bad(20, f"Batch!B{r}", "COMPANY is empty", value="(empty)")
            k = _norm(batch.cell(r, 5).value)
            if k:
                db_rows.setdefault(k, r)      # the same DB# on several rows is allowed: the amounts are summed and matched to the Ageing (P26, P3)

    # 26 (AGEING) ---- each Batch row against the Ageing: DB# -> Ref#, AMOUNT = Total, COMPANY = Name, CLAIM # = Claim#
    db_group = {}             # normalised DB# -> amounts of every Batch row with that DB#
    memo_group = {}           # normalised DB# -> [(sheet, memo total)] for DB#s repeated on the Batch
    if batch and filled_rows:
        if not ref["ageing_file"]:
            bad(26, "Batch!E8", "No Ageing file imported yet. Upload it in the DataBase Room.", value="(no Ageing)")
        else:
            memo_who = {int(n): (_norm(wb[n].cell(6, 2).value), _txt(wb[n].cell(4, 2).value).strip()) for n in memos}
            for pos, r in enumerate(filled_rows, start=1):
                k = _norm(batch.cell(r, 5).value)
                if k:
                    v = batch.cell(r, 4).value
                    db_group.setdefault(k, []).append(v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0)
            seen_group = set()
            for pos, r in enumerate(filled_rows, start=1):
                db = _txt(batch.cell(r, 5).value).strip()
                if not db:
                    continue                                    # an empty DB# is reported by the DB# check
                acct_k, name_k = memo_who.get(pos, ("", ""))   # Batch row k belongs to memo sheet k: its SOLD TO account and name filter the Ageing
                hit, anywhere = _ageing_find(ref, db, acct_k, name_k)
                if not hit:
                    bad(26, f"Batch!E{r}", _not_found(db, acct_k, anywhere), value=db, expected=acct_k or None)
                    continue
                a = hit[0]
                amt = _num(batch.cell(r, 4).value)
                grp = db_group.get(_norm(db), [])
                if len(grp) > 1:      # repeated DB#: group by DB#, sum the AMOUNT and match the Ageing Total (reported once, on the first row)
                    if _norm(db) not in seen_group:
                        seen_group.add(_norm(db))
                        tot = round(sum(grp), 2)
                        if a["total"] is not None and abs(tot - a["total"]) > MONEY_TOL:
                            bad(26, f"Batch!D{r}", f"{db} is on {len(grp)} rows: they add up to {tot:,.2f} but the Ageing Total is {a['total']:,.2f}", "Make the rows of this DB# add up to the Ageing Total", value=f"{tot:,.2f} vs {a['total']:,.2f}", expected=f"{a['total']:,.2f}")
                elif amt is not None and a["total"] is not None and abs(amt - a["total"]) > MONEY_TOL:
                    bad(26, f"Batch!D{r}", f"Amount {amt:,.2f} but the Ageing Total of {db} is {a['total']:,.2f}", value=f"{amt:,.2f} vs {a['total']:,.2f}", expected=f"{a['total']:,.2f}")
                company = _txt(batch.cell(r, 2).value).strip()
                if _norm(company) != _norm(a["name"]):
                    bad(26, f"Batch!B{r}", f"COMPANY \"{company}\" but the Ageing Name of {db} is \"{_txt(a['name']).strip()}\"", value=company or "(empty)", expected=_txt(a["name"]).strip())
                claim_b, claim_a = _txt(batch.cell(r, 6).value).strip(), _txt(a["claim_no"]).strip()
                if claim_b != claim_a:
                    bad(26, f"Batch!F{r}", f"CLAIM # \"{claim_b}\" but the Ageing Claim# of {db} is \"{claim_a}\"", value=claim_b or "(empty)", expected=claim_a or None)

    # 15 (EMPTY) RETURN on the Batch sheet; 16 (WORK) work sheets
    if batch and _txt(batch["D3"].value).strip():
        bad(15, "Batch!D3", "RETURN must be empty", value=_txt(batch["D3"].value))
    for s in ("Work1", "Work2", "Work3"):
        if s in wb.sheetnames:
            for row in wb[s].iter_rows():
                for c in row:
                    if c.value is not None and _txt(c.value).strip():
                        bad(16, f"{s}!{c.coordinate}", f"{s} must be empty", value=c.value)

    # 18 (MATH) Batch rows: amounts are numbers and the total is their sum
    if batch and filled_rows:
        amounts = []
        for r in filled_rows:
            v = batch.cell(r, 4).value
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                bad(18, f"Batch!D{r}", "The amount must be a number", value=_txt(v) or "(empty)")
            else:
                amounts.append(v)
                if v <= 0:
                    bad(22, f"Batch!D{r}", "The amount must be more than zero", value=f"{v:,.2f}")
        if total_cell is not None:
            shown = vals["Batch"].cell(total_cell, 4).value
            if isinstance(shown, (int, float)) and abs(shown - sum(amounts)) > MONEY_TOL:
                bad(18, f"Batch!D{total_cell}", f"The Batch total is {shown:,.2f} but the rows add up to {sum(amounts):,.2f}", value=f"{shown:,.2f} vs {sum(amounts):,.2f}")
            elif unnumbered and isinstance(shown, (int, float)):       # only the numbered rows count: a row without a number is left out of the sum
                numbered = sum(v for r in filled_rows if r not in {u for u, _ in unnumbered} for v in [batch.cell(r, 4).value] if isinstance(v, (int, float)))
                if abs(shown - numbered) > MONEY_TOL:
                    bad(18, f"Batch!D{total_cell}", f"The Batch total is {shown:,.2f} but the numbered rows add up to {numbered:,.2f} (row {unnumbered[0][0]} has no number)", value=f"{shown:,.2f} vs {numbered:,.2f}")
    dates = []                # (where, value) of every date cell, for the SAME DATE check
    if batch:
        dates += [(f"Batch!C{r}", batch.cell(r, 3).value) for r in filled_rows]
    checked = 0
    for n in memos:
        ws, wv = wb[n], vals[n]
        g = lambda r, c: ws.cell(r, c).value
        sub = next((r for r in range(1, ws.max_row + 1) if _norm(g(r, 5)) == "SUBTOTAL"), None)
        if sub is None:
            continue
        lines = [(r, _txt(g(r, 1)).strip(), _num(g(r, 4)), _num(g(r, 7))) for r in range(13, sub) if _txt(g(r, 1)).strip()]
        sold_to, sold_acct, ship_acct = _txt(g(4, 2)).strip(), _norm(g(6, 2)), _norm(g(6, 7))
        if not lines and not sold_to and not sold_acct:
            continue                                     # empty form (the template's sample sheet)
        checked += 1
        sheet = f"Memo {n}"
        tot_row = next((r for r in range(sub, ws.max_row + 1) if _norm(g(r, 7)) == "TOTAL"), None)
        db_row = next((r for r in range(1, ws.max_row + 1) if _norm(g(r, 1)) == "DEBIT MEMO NOS."), None)
        memo_db = _txt(g(db_row + 1, 1)).strip() if db_row else ""
        b = batch_rows.get(int(n))

        # 2 ---- hashtag, same number in both places
        for where, val in ((f"Batch!E{b['row']}", b["db"].strip()) if b else (None, None), (f"{sheet}!A{(db_row or 0) + 1}", memo_db)):
            if where is None:
                continue
            if not DB_RE.match(val):
                bad(2, where, f"\"{val}\" is not like DB#123456" if val else "DB# is empty", value=val or "(empty)")
        if b is None:
            bad(2, sheet, f"No row {n} on the Batch sheet for this memo")
        elif b["db"].strip() != memo_db:
            bad(2, sheet, f"Batch has \"{b['db'].strip()}\" but the memo has \"{memo_db}\"", value=f"{b['db'].strip()} != {memo_db}", expected=memo_db or None)

        # 3 ---- DB amount from Ageing
        memo_total = wv.cell(tot_row, 9).value if tot_row else None
        if not isinstance(memo_total, (int, float)):
            memo_total = round(sum((q or 0) * (p or 0) for _, _, q, p in lines) + sum(_num(g(r, 9)) or 0 for r in range(sub + 1, (tot_row or sub + 4))), 2)
        hit, anywhere = _ageing_find(ref, memo_db, sold_acct, sold_to)
        if not ref["ageing_file"]:
            bad(3, sheet, "No Ageing file imported yet. Upload it in the DataBase Room.")
        elif not hit:
            bad(3, sheet, _not_found(memo_db, sold_acct, anywhere) if memo_db else "DB# is empty", value=memo_db or "(empty)", expected=sold_acct or None)
        else:
            amt = hit[0]["total"]
            if len(db_group.get(_norm(memo_db), [])) > 1:     # repeated DB#: the memos of that DB# are summed together (checked after the loop)
                memo_group.setdefault(_norm(memo_db), {"db": memo_db, "amt": amt, "sheets": [], "sum": 0.0})
                memo_group[_norm(memo_db)]["sheets"].append(n)
                memo_group[_norm(memo_db)]["sum"] += memo_total or 0
            elif abs((memo_total or 0) - (amt or 0)) > 0.005:
                bad(3, sheet, f"Memo total {memo_total:,.2f} but Ageing has {amt:,.2f} for {memo_db}", value=f"{memo_total:,.2f} vs {amt:,.2f}")
            if len(db_group.get(_norm(memo_db), [])) > 1:
                pass
            elif b and b["amount"] is not None and abs(b["amount"] - (amt or 0)) > 0.005:
                bad(3, f"Batch!D{b['row']}", f"Batch amount {b['amount']:,.2f} but Ageing has {amt:,.2f} for {memo_db}", value=f"{b['amount']:,.2f} vs {amt:,.2f}")

        # 4 + 10 ---- sale report of the same account
        if not _has_sale_report(ref, sold_acct, ship_acct):
            note(f"{sheet}!A13", f"No Sale Report of account {sold_acct or '(no account)'} is uploaded, so the items, quantities and prices were not checked", "Upload the Sale Report of this account, then check again", None, "Sale Report not uploaded")
        else:
            for r, item, qty, price in lines:
                sl = _sale_lines(ref, sold_acct, ship_acct, _norm(item))
                if not sl:
                    bad(4, f"{sheet}!A{r}", f"{item} is not in the sale report of {sold_acct}", value=item)
                    continue
                if qty is not None and qty > sum(x[1] for x in sl) + 1e-9:
                    bad(10, f"{sheet}!D{r}", f"{item}: quantity {qty:g} is more than the sale report ({sum(x[1] for x in sl):g})", value=f"{qty:g} vs {sum(x[1] for x in sl):g}")
                top = max((x[2] for x in sl if x[2] is not None), default=None)
                if price is not None and top is not None and price > top + 1e-9:
                    bad(10, f"{sheet}!G{r}", f"{item}: price {price:g} is more than the sale report ({top:g})", value=f"{price:g} vs {top:g}")

        # 5 ---- salesman, terms, reason, reason code
        reason, code, slm, terms = _txt(g(9, 1)).strip(), _txt(g(9, 3)).strip(), _txt(g(9, 5)).strip(), _txt(g(9, 6)).strip()
        # 34 ---- reason codes that are not allowed on a credit memo
        if is_blocked(code, reason):
            bad(34, f"{sheet}!C9", f"Reason code {code} ({reason}) is not allowed on a credit memo", value=f"{code} · {reason}")
        for label, v, cell in (("Reason", reason, "A9"), ("Reason code", code, "C9"), ("Salesman", slm, "E9"), ("Terms", terms, "F9")):
            if not v:
                bad(5, f"{sheet}!{cell}", f"{label} is empty", value="(empty)")
        # the memo reason must be the same, case sensitive, as in the REASON sheet (checked once that sheet matches the database)
        if code and sheet_ok:
            db_desc = (ref["reasons"] or {}).get(_norm(code))
            if code not in reasons and db_desc is not None:      # second priority: the Reason codes list of the DataBase Room
                if reason and reason != db_desc:
                    bad(5, f"{sheet}!A9", f"Reason must be exactly \"{db_desc}\" for code {code} (Reason codes database)", value=reason, expected=db_desc)
            elif code not in reasons:
                near = next((c for c in reasons if c.upper() == code.upper()), None)
                bad(5, f"{sheet}!C9", f"Reason code must be exactly {near} (case sensitive)" if near else f"Reason code {code} is not in the REASON sheet nor in the Reason codes database",
                    f"Write the code as {near}" if near else None, value=code)
            elif reason and reason != reasons[code]:
                bad(5, f"{sheet}!A9", f"Reason must be exactly \"{reasons[code]}\" for code {code} (case sensitive)", value=reason, expected=reasons[code])
        an = ref["analyst"].get(sold_acct)
        if an:
            if slm and not _slm_match(slm, an["slm"]):
                bad(5, f"{sheet}!E9", f"Salesman \"{slm}\" but the account has \"{an['slm']}\"", value=slm, expected=an["slm"])
        # Terms: build the Cust Code (Batch DIVISION D4 + SOLD TO account B6) and verify the Terms against Report 2
        cc = sold_acct if batch else ""
        cc_rows = ref["cust"].get(cc) if sold_acct else None
        if terms and cc_rows and not any(_norm(terms) in x[0] for x in cc_rows):
            want = sorted(x[0] for x in cc_rows)[0]
            bad(5, f"{sheet}!F9", f"Terms \"{terms}\" but Report 2 has \"{want}\" for Cust Code {cc}", value=terms, expected=want)

        # 20 (REQUIRED) SOLD TO;  21 (CAPS) accounts in capitals;  22 (POSITIVE);  24 (GAPS)
        if not sold_to:
            bad(20, f"{sheet}!B4", "SOLD TO name is empty", value="(empty)")
        if not sold_acct:
            bad(20, f"{sheet}!B6", "SOLD TO account is empty", value="(empty)")
        for cell, raw in ((f"{sheet}!B6", _txt(g(6, 2)).strip()), (f"{sheet}!G6", _txt(g(6, 7)).strip())):
            if raw and (raw != raw.upper() or re.search(r"\s", raw)):
                bad(21, cell, f"Account \"{raw}\" must be in capitals with no spaces ({re.sub(r'\s+', '', raw).upper()})", value=raw, expected=re.sub(r"\s+", "", raw).upper())
        for r, item, q, p in lines:
            if item != item.upper():
                note(f"{sheet}!A{r}", f"Style no. \"{item}\" is not in capitals", "Write it in capitals", item, "Style number not in capitals")
            if q is not None and (q <= 0 or q != int(q)):
                bad(22, f"{sheet}!D{r}", f"{item}: PCS must be a whole number above zero", value=f"{q:g}")
            if p is not None and p <= 0:
                bad(22, f"{sheet}!G{r}", f"{item}: price must be above zero", value=f"{p:g}")
        for r in range(sub + 1, (tot_row or sub + 4)):
            v = g(r, 9)
            if v not in (None, "") and _num(wv.cell(r, 9).value) is None:
                bad(22, f"{sheet}!I{r}", f"{_txt(g(r, 5)).strip() or 'This cell'} must be a number", value=_txt(v))
        blank_at = None
        for r in range(13, sub):
            if not _txt(g(r, 1)).strip():
                blank_at = blank_at or r
            elif blank_at:
                bad(24, f"{sheet}!A{blank_at}", f"Empty line between items (row {blank_at}); the next item is on row {r}", value="(empty)")
                break

        # 14 (CLAIM) ---- CUSTOMER CLAIM # is "<Reason> : <claim no.>" or just "<Reason>"
        claim = _txt(g(10, 5)).strip()
        claim_no = None
        if reason:
            claim_no = _claim_no(claim, reason)
            if claim_no is None:
                bc = _txt(batch.cell(b["row"], 6).value).strip() if b else ""
                want = f"{reason} : {bc}" if bc else reason
                spacing = claim.startswith(reason) and re.match(r"\s*:\s*\S", claim[len(reason):]) is not None
                msg = (f"Spacing around \":\" must be exactly one space each side: \"{reason} : <claim no.>\"" if spacing
                       else f"Claim # must be \"{reason} : <claim no.>\" or just \"{reason}\"")
                bad(14, f"{sheet}!E10", msg, f"Write it as \"{want}\"", value=claim or "(empty)", expected=want)
        # 15 (EMPTY) ---- SHIPPED VIA
        if _txt(g(9, 8)).strip():
            bad(15, f"{sheet}!H9", "SHIPPED VIA must be empty", value=_txt(g(9, 8)))
        # 17 (DATE) ---- collected, compared after the loop
        dates.append((f"{sheet}!H2", g(2, 8)))

        # 18 (MATH) ---- lines, SUBTOTAL, TOTAL, Batch vs memo
        for r in range(13, sub):
            style, q, p = _txt(g(r, 1)).strip(), g(r, 4), g(r, 7)
            if not style and (_txt(q).strip() or _txt(p).strip()):
                bad(18, f"{sheet}!A{r}", "A line with PCS or price must have a style no.", value="(empty)")
            if not style:
                continue
            if _num(q) is None:
                bad(18, f"{sheet}!D{r}", f"{style}: PCS must be a number", value=_txt(q) or "(empty)")
            if _num(p) is None:
                bad(18, f"{sheet}!G{r}", f"{style}: price must be a number", value=_txt(p) or "(empty)")
            if _txt(wv.cell(r, 5).value).strip() == "#N/A":
                bad(18, f"{sheet}!E{r}", f"{style} is not on the PRICE LIST (description shows #N/A)", "Add the style to the PRICE LIST", value="#N/A")
            shown = wv.cell(r, 9).value
            if _num(q) is not None and _num(p) is not None and isinstance(shown, (int, float)) and abs(shown - _num(q) * _num(p)) > MONEY_TOL:
                bad(18, f"{sheet}!I{r}", f"{style}: amount {shown:,.2f} but PCS x price is {_num(q) * _num(p):,.2f}", value=f"{shown:,.2f} vs {_num(q) * _num(p):,.2f}")
        line_sum = sum((_num(g(r, 4)) or 0) * (_num(g(r, 7)) or 0) for r in range(13, sub) if _txt(g(r, 1)).strip())
        sub_shown = wv.cell(sub, 9).value
        if isinstance(sub_shown, (int, float)) and abs(sub_shown - line_sum) > MONEY_TOL:
            bad(18, f"{sheet}!I{sub}", f"SUBTOTAL is {sub_shown:,.2f} but the lines add up to {line_sum:,.2f}", value=f"{sub_shown:,.2f} vs {line_sum:,.2f}")
        if tot_row is not None and isinstance(sub_shown, (int, float)):
            extras = sum(_num(wv.cell(r, 9).value) or 0 for r in range(sub + 1, tot_row))
            tot_shown = wv.cell(tot_row, 9).value
            if isinstance(tot_shown, (int, float)) and abs(tot_shown - (sub_shown + extras)) > MONEY_TOL:
                bad(18, f"{sheet}!I{tot_row}", f"TOTAL is {tot_shown:,.2f} but SUBTOTAL + GST/HST + FREIGHT/HANDLING + ADJUSTMENT is {sub_shown + extras:,.2f}", value=f"{tot_shown:,.2f} vs {sub_shown + extras:,.2f}")
        if b:
            if b["amount"] is not None and isinstance(memo_total, (int, float)) and abs(b["amount"] - memo_total) > MONEY_TOL:
                bad(18, f"Batch!D{b['row']}", f"Batch amount {b['amount']:,.2f} but the memo TOTAL is {memo_total:,.2f}", value=f"{b['amount']:,.2f} vs {memo_total:,.2f}")
            batch_claim = _txt(batch.cell(b["row"], 6).value).strip()
            if claim_no is not None and batch_claim != claim_no:
                bad(18, f"Batch!F{b['row']}", f"Batch CLAIM # \"{batch_claim}\" but the memo has \"{claim_no}\"", value=batch_claim or "(empty)", expected=claim_no or None)
            if sold_to and _norm(b["company"]) != _norm(sold_to):
                note(f"Batch!B{b['row']}", f"COMPANY \"{b['company'].strip()}\" differs from SOLD TO \"{sold_to}\"", None, b["company"].strip(), "Batch COMPANY differs from SOLD TO")

        # 7 ---- template formulas must stay
        for r, item, _, _ in lines:
            if g(r, 5) != f"=VLOOKUP(A{r},'PRICE LIST'!A:B,2,0)":
                bad(7, f"{sheet}!E{r}", "VLOOKUP formula is missing or changed" if str(g(r, 5) or "").startswith("=") else f"No formula =VLOOKUP(A{r},'PRICE LIST'!A:B,2,0) found. Manual value {_txt(g(r, 5)).strip() or '(empty)'} was typed", f"Restore =VLOOKUP(A{r},'PRICE LIST'!A:B,2,0)", value=g(r, 5))
            if not _line_formula(g(r, 9), r):
                cur = g(r, 9)
                if isinstance(cur, str) and cur.startswith("="):
                    msg = f"Template formula was changed to {cur}"
                else:
                    msg = f"No formula =G{r}*D{r} found. Manual value {_txt(cur).strip() or '(empty)'} was typed"
                bad(7, f"{sheet}!I{r}", msg, f"Restore =G{r}*D{r}", value=cur)
        if not re.fullmatch(r"=SUM\(I13:I\d+\)", _txt(g(sub, 9))):
            bad(7, f"{sheet}!I{sub}", "SUBTOTAL formula is missing or changed", value=g(sub, 9))
        if tot_row is None or not re.fullmatch(r"=SUM\(I\d+:I\d+\)", _txt(g(tot_row, 9))):
            bad(7, f"{sheet}!I{tot_row or sub + 4}", "TOTAL formula is missing or changed", value=g(tot_row, 9) if tot_row else "(empty)")

        # 8 ---- account names from Ageing
        ag_name = ref["names"].get(sold_acct)
        if not ref["ageing_file"]:
            pass
        elif ag_name is None:
            bad(8, f"{sheet}!B6", f"Account {sold_acct or '(empty)'} is not in the Ageing", value=sold_acct or "(empty)")
        else:
            if _norm(sold_to) != _norm(ag_name):
                bad(8, f"{sheet}!B4", f"SOLD TO \"{sold_to}\" but Ageing has \"{ag_name}\"", value=sold_to, expected=ag_name)
            if b and _norm(b["company"]) != _norm(ag_name):
                bad(8, f"Batch!B{b['row']}", f"COMPANY \"{b['company'].strip()}\" but Ageing has \"{ag_name}\"", value=b["company"].strip(), expected=ag_name)

        # 9 ---- duplicate item in one memo
        seen = {}
        for r, item, _, price in lines:       # the same item twice is fine with a different price; an error only with the same price too
            prev = seen.get(_norm(item), [])
            same = next((pr for pr, pp in prev if pp == price or (pp is not None and price is not None and abs(pp - price) < 1e-9)), None)
            if same is not None:
                bad(9, f"{sheet}!A{r}", f"{item} is also on row {same} with the same price", value=item)
            seen.setdefault(_norm(item), []).append((r, price))

    # 3 ---- repeated DB#: the memos of that DB# add up to the Ageing amount
    for gk, gm in memo_group.items():
        tot = round(gm["sum"], 2)
        if abs(tot - (gm["amt"] or 0)) > 0.005:
            bad(3, f"Memo {gm['sheets'][0]}", f"{gm['db']} is on memos {', '.join(gm['sheets'])}: they add up to {tot:,.2f} but Ageing has {gm['amt']:,.2f}",
                "Make the memos of this DB# add up to the Ageing amount", value=f"{tot:,.2f} vs {gm['amt']:,.2f}", expected=f"{gm['amt']:,.2f}")

    # same style no. with a different price on another memo sheet: a note only, never an error
    first_price = {}
    for n in memos:
        w = wb[n]
        end = next((r for r in range(1, w.max_row + 1) if _norm(w.cell(r, 5).value) == "SUBTOTAL"), None)
        for r in range(13, end or 13):
            st, pr = _norm(w.cell(r, 1).value), _num(w.cell(r, 7).value)
            if not st or pr is None:
                continue
            if st not in first_price:
                first_price[st] = (n, pr)
            elif abs(first_price[st][1] - pr) > 1e-9:
                note(f"Memo {n}!G{r}", f"{_txt(w.cell(r, 1).value).strip()}: price {pr:g} is different from memo {first_price[st][0]} ({first_price[st][1]:g})",
                     "Check that the price is right", f"{pr:g} vs {first_price[st][1]:g}", "Same style, different price")

    # 29 (MEMOCOMPANY) ---- memo SOLD TO name (B4) = the COMPANY of its own Batch row (memo n = Batch row n) = Ageing Name of its SOLD TO account (B6); one batch may hold several companies
    if batch:
        for n in memos:
            got = _txt(wb[n]["B4"].value).strip()
            ag = ref["names"].get(_norm(wb[n]["B6"].value)) if ref["ageing_file"] else None
            if not got:
                continue
            company = _txt(batch.cell(7 + int(n), 2).value).strip() if str(n).isdigit() else ""
            if company and _norm(got) != _norm(company):
                bad(29, f"Memo {n}!B4", f"SOLD TO \"{got}\" is different from the Batch COMPANY of row {n} (\"{company}\")", value=got, expected=company)
            elif ag and _norm(got) != _norm(ag):
                bad(29, f"Memo {n}!B4", f"SOLD TO \"{got}\" is different from the Ageing Name (\"{_txt(ag).strip()}\")", value=got, expected=_txt(ag).strip())

    # 30 (SHIPCODE) ---- memo SHIP TO account (G6) is an Acct ShipCode of the Sale Report where Dv+SaleReport = Batch D4 + memo B6; else a Ship Code of Report 2
    if batch:
        for n in memos:
            ship, key = _norm(wb[n]["G6"].value), _norm(batch["D4"].value) + _norm(wb[n]["B6"].value)
            if not ship or key == _norm(batch["D4"].value):
                continue                      # empty cells are reported by the Required / Same account protocols
            if not _has_sale_report(ref, _norm(wb[n]["B6"].value), ship):
                continue                      # no Sale Report of this account uploaded: nothing to check against (the note is on protocol 4)
            if ship not in ref["sale_ship"].get(key, ()) and ship not in ref["ar_ship"].get(key, ()):
                bad(30, f"Memo {n}!G6", f"Ship code {ship} was not found in the Sale Report nor in Report 2 for {key}", value=ship, expected=None)

    # 31 (CUSTCODE) ---- Cust Code of Report 2 = Batch D4 + memo B6; memo SALESMAN (E9) and TERMS (F9) match its SLM and Terms
    if batch and ref["cust"]:
        for n in memos:
            key = _norm(wb[n]["B6"].value)
            slm, terms = _norm(wb[n]["E9"].value), _norm(wb[n]["F9"].value)
            if not key:
                continue                      # empty account: reported by the Required protocol
            rows = ref["cust"].get(key)
            if not rows:
                bad(31, f"Memo {n}!B6", f"Cust Code {key} was not found in Report 2", value=key)
                continue
            if slm and not any(_slm_match(slm, x[1]) for x in rows):
                bad(31, f"Memo {n}!E9", f"Salesman \"{_txt(wb[n]['E9'].value).strip()}\" but Report 2 has \"{sorted(x[1] for x in rows)[0]}\" for {key}", value=_txt(wb[n]["E9"].value).strip(), expected=sorted(x[1] for x in rows)[0])

    # 17 (DATE) ---- every date is filled and equals the batch DATE (D2); with D2 empty the first date found is used (and a note says so)
    ref_date = _date(batch["D2"].value) if batch else None
    if batch and ref_date is None and not _txt(batch["D2"].value).strip():
        first = next(((w, _date(v)) for w, v in dates if _date(v)), None)
        if first:
            ref_date = first[1]
            note("Batch!D2", f"The batch DATE is empty; {first[1]:%m/%d/%Y} (first date found) is used", "Fill in the batch DATE", None, "Batch DATE is empty")
    for where, v in dates:
        d = _date(v)
        if d is None:
            bad(17, where, "The date is empty or not a date", value=_txt(v) or "(empty)")
        elif ref_date and d != ref_date:
            bad(17, where, f"{d:%m/%d/%Y} but the batch date is {ref_date:%m/%d/%Y}", value=f"{d:%m/%d/%Y}", expected=f"{ref_date:%m/%d/%Y}")

    # Batch rows with no memo sheet
    for k, b in batch_rows.items():
        if str(k) not in memos and any(batch.cell(b["row"], c).value not in (None, "") for c in range(2, 11)):   # empty template rows are fine
            bad(2, f"Batch!A{b['row']}", f"Batch row {k} has no memo sheet", value=k)

    results = [{"n": n, "label": TAGS.get(n, f"P{n}") if n in (19, 32) else f"P{n}", "tag": TAGS.get(n, f"P{n}"), "title": t, "info": i, "pass": not issues[n], "issues": issues[n][:200], "count": len(issues[n])}
               for n, t, i in CHECKS]
    # stages: Batch format, then Batch values + Ageing, then everything else; a later stage is only shown when the earlier ones are clean
    by_n = {r["n"]: r for r in results}
    staged = [ns for _, ns in STAGES]
    staged.append([n for n, *_ in CHECKS if not any(n in ns for ns in staged)])
    stopped = None
    for k, ns in enumerate(staged):
        for n in ns:
            by_n[n]["stage"] = k                 # the viewer lists the errors in this order
    for (name, _), ns in zip(STAGES + [("The other checks", [])], staged):
        if not STOP_AT_FAILED_STAGE:
            break
        if stopped:
            for n in ns:
                by_n[n].update(issues=[], count=0, skipped=True, **{"pass": False})
        elif any(by_n[n]["count"] for n in ns):
            stopped = name
    total = sum(b["amount"] or 0 for b in batch_rows.values())
    bdate = batch["D2"].value if batch else None
    return {"total": round(total, 2), "date": bdate.strftime("%m/%d/%Y") if hasattr(bdate, "strftime") else None, "sheets": wb.sheetnames, "memos": checked, "ageing_file": ref["ageing_file"], "sale_files": _memo_sale_files(wb, memos, ref),
            "passed": sum(r["pass"] for r in results), "results": results, "notes": notes[:200],
            "stopped": stopped if stopped != "The other checks" else None, "skipped": sum(1 for r in results if r.get("skipped"))}
