"""The strict rules of the user's Credit Memo Batch template (templates_xlsx/Credit Memo Batch TEMPLATE.xlsx).

Exports and validation may only use these sheets, columns and formulas. Nothing else is allowed.
"""

TEMPLATE_FILE = "Credit Memo Batch TEMPLATE.xlsx"

SHEETS = ["REASON", "PRICE LIST", "Work1", "Work2", "Work3", "Batch", "1"]  # "1" is copied once per memo

BATCH_COLUMNS = ["#", "COMPANY", "DATE", "AMOUNT", "DB#", "CLAIM #", "RT #", "RT AMOUNT", "DISCREPANCY", "CREDIT OVER / UNDER"]
BATCH_FIRST_ROW, BATCH_LAST_ROW = 8, 27  # 20 numbered rows in the template
BATCH_HEADER_CELLS = {"batch_no": "D1", "date": "D2", "return": "D3", "division": "D4"}

MEMO_FIRST_LINE, MEMO_LAST_LINE = 13, 33  # 21 item rows in the template
MEMO_LINE_COLUMNS = {"A": "STYLE NO. (SKU)", "D": "QUANTITY (PCS)", "E": "description (formula)", "G": "PRICE", "I": "AMOUNT (formula)"}

# Only these formulas may appear in a generated file ({r} = row number, {n} = row numbers of the memo).
ALLOWED_FORMULAS = {
    "batch_total": "=SUM(D8:D27)",
    "memo_description": "=VLOOKUP(A{r},'PRICE LIST'!A:B,2,0)",
    "memo_amount": "=G{r}*D{r}",
    "memo_subtotal": "=SUM(I13:I33)",
    "memo_total": "=SUM(I34:I37)",
}
