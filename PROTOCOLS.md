# Recheck protocols

Source of truth: `CHECKS` in `app/recheck.py`. Keep this file in step with it: when a protocol is added or changed, update this file and its test in `tests/test_recheck.py`.

| # | Protocol | Rule | How to fix | Shown as |
|---|---|---|---|---|
| 1 | No spaces in cells | No space before or after the text in any cell, on every sheet. A cell with only a space is also wrong. | Remove the extra space (press Delete) | P1 |
| 2 | DB# with hashtag | DB#123456 (or IN#123456). The # is a must. Same number on the Batch sheet and the memo sheet. | Write it with the hashtag, like DB#123456, the same on Batch and memo | P2 |
| 3 | DB amount = Ageing | The DB number must exist in the Ageing and the memo amount must equal the Ageing amount. A DB# on several memos: the memo totals are added and must equal the Ageing amount. | Check the DB in the Ageing first; the memo amount must be the same | P3 |
| 4 | Item in sale report | Every item on the memo must be in the sale report of the same account.  If no Sale Report of the account is uploaded this is not checked; a note asks to upload it. | Use an item that was sold to this account | P4 |
| 5 | Salesman, Terms, Reason | All four filled. Salesman matches the account: the names are split on / & , ; and at least one name must be the same (JACK/AARON matches AARON; case and extra spaces ignored). Terms: the Cust Code is built from the Batch DIVISION (D4) + SOLD TO account (B6) (EM + AM9286 = EMAM9286) and the Terms must match its Terms in Report 2. Reason and code are exactly (case sensitive) as in the REASON sheet; a code that is not there is looked up in the Reason codes list of the DataBase Room. | Fill it in; use the code and exact text from the REASON sheet | P5 |
| 6 | No extra formula | Only the Batch total cell has a formula on the Batch sheet. Nothing else may be added anywhere. | Delete the extra formula and type the value by hand | P6 |
| 7 | Template formulas kept | VLOOKUP, price x quantity (=G13*D13 or =D13*G13, both accepted), SUBTOTAL and TOTAL must stay on every memo. A typed number instead of the formula shows "No formula =G13*D13 found. Manual value 120 was typed". | Restore the template formula | P7 |
| 8 | Account name = Ageing | SOLD TO name and account, and the Batch COMPANY, must be the same as in the Ageing. | Copy the account name and number from the Ageing | P8 |
| 9 | No duplicate item | One item once on the PRICE LIST. On a memo the same item twice is an error only when the price is the same too. | Keep each item only once | P9 |
| 10 | Qty and price within sale report | Memo quantity and price must not be more than the sale report. Less is allowed.  If no Sale Report of the account is uploaded this is not checked; a note asks to upload it. | Use a quantity and price not more than the sale report | P10 |
| 11 | REASON sheet = database | The REASON sheet of the file must match the Reason codes list imported in the DataBase Room (case does not matter). Then every memo reason and code must match that REASON sheet exactly (case sensitive). | Make the REASON sheet the same as the Reason codes database | D_P11 |
| 12 | Format unchanged | The template sheets, labels and layout are intact: required sheets, no extra sheets, Batch and memo labels, Batch total formula. A column width that differs from the template is only a note. | Use the template as it is; do not rename, delete or add sheets or labels | FORMAT |
| 13 | Memo sheets in sequence | At least 1 memo sheet (more than 20 allowed), numbered 1, 2, 3 without gaps. Batch rows numbered 1, 2, 3 and as many as memo sheets. | Number the memo sheets 1, 2, 3... and keep one Batch row per memo | SEQ |
| 14 | Claim # format | CUSTOMER CLAIM # is '<Reason> : <claim no.>', or just '<Reason>' when there is no claim number. The claim no. is the Batch CLAIM # of that memo (any text, numbers, symbols); one space each side of the colon. | Write it like: Reason : claim number (or only the Reason) | CLAIM |
| 15 | Ship via / Return empty | The RETURN cell on the Batch sheet and SHIPPED VIA on every memo must be empty. | Clear the cell | EMPTY |
| 16 | Work sheets empty | Work1, Work2 and Work3 must have no data. | Clear the Work sheets | WORK |
| 17 | Same date | Every Batch row date and every memo date is filled and equals the batch DATE. | Use the batch date in every date cell | DATE |
| 18 | Math and cross-checks | Line amount = PCS x price, SUBTOTAL, TOTAL, Batch total, and Batch row = memo (amount, claim #). | Correct the amount so the math agrees | MATH |
| 19a | Division in capitals | The DIVISION on the Batch sheet is two letters, both in capitals. | Type the division with both letters in capitals | P19a |
| 20 | Required fields filled | DATE, COMPANY on every Batch row, and SOLD TO name and account on every memo are filled in (the batch number may be empty; if written in another format than 10-26973 it is only a warning). | Fill it in | REQUIRED |
| 21 | Account in capitals | SOLD TO and SHIP TO account numbers are written in capitals with no spaces (AM9286). Style numbers in small letters show as a note. | Write the account in capitals without spaces | CAPS |
| 22 | Positive numbers | PCS is a whole number above zero, price and amounts are above zero, and GST/HST, freight and adjustment are numbers. | Use a whole number above zero for PCS and an amount above zero | POSITIVE |
| 23 | No duplicate code | A reason code may be on the REASON sheet only once. A DB# may repeat on the Batch: its rows are grouped (see 26 and 3). | Keep each reason code only once | DUPLICATE |
| 24 | No empty line between items | On a memo the items follow each other with no empty line between them. | Move the items up so there is no empty line between them | GAPS |
| 25 | PRICE LIST complete | Every item on the PRICE LIST has a description. | Add the description of the item on the PRICE LIST | PRICELIST |
| 26 | Batch = Ageing | Every Batch row: DB# is in the Ageing Ref#; AMOUNT = Ageing Total (a DB# repeated on several rows: group by DB#, SUM(AMOUNT) = Ageing Total, reported once); COMPANY = Ageing Name of that DB# (case does not matter); CLAIM # = Ageing Claim#. (The DATE must be the batch DATE: see Same date.) | Copy the value from the Ageing | AGEING |
| 29 | Memo company = Batch company = Ageing | SOLD TO name (B4) of every memo is the same as the COMPANY of its own Batch row (memo 3 = Batch row 3) and as the Ageing Name of its account (case does not matter). One batch may hold several companies, so memos and rows are not compared with each other. | Copy the Ageing name into SOLD TO on every memo and into the Batch COMPANY | MEMOCOMPANY |
| 30 | Ship code in Sale Report / Report 2 | SHIP TO account (G6) on every memo must be an Acct ShipCode of the Sale Report, or else a Ship Code of Report 2 (A/R PDF), for the account Batch DIVISION (D4) + SOLD TO account (B6).  If no Sale Report of the account is uploaded this is not checked; a note asks to upload it. | Use a ship code that is on the Sale Report or Report 2 for this account | SHIPCODE |
| 31 | Salesman = Report 2 | The Report 2 Cust Code is read as: text before the first space, first 2 letters skipped (`EMMR994450 MRL` -> MR994450); it must equal the SOLD TO account (B6) (the Report 2 value is stored as it is); the memo SALESMAN (E9) must match its SLM by name, at least one name the same, the first word of the Report salesman is skipped and the rest is compared (AAA ISAAC ASH matches ISAAC ASH; J35 JACK/AARON matches JACK or AARON) (the Terms are checked by Salesman, Terms, Reason). | Copy the salesman and terms of the Cust Code from Report 2 | CUSTCODE |
| 32 | Division = Ageing Div | The Batch DIVISION (D4) must be a Div of the Ageing for the SOLD TO account (B6) of the memos. The Ageing is grouped by Acct and Div; if D4 is one of them it is ok. | Use the Div of the Ageing for this account in the Batch DIVISION | P19 |
| 33 | PRICE LIST = memos | The PRICE LIST has exactly the items of the memo sheets (1 to 20, and 21 if present): every memo item must be on the PRICE LIST, and any item that is on no memo is extra and must be removed. Items are compared ignoring case and spaces. | Add the missing memo items and delete the extra items from the PRICE LIST | PLMEMOS |
| 14 (note) | Claim # with DNT | When the claim no. contains DNT, the claim may start with a percentage: `2% Damage Allowance : DNT1655357`. | | CLAIM |
| 34 | Reason not allowed | No memo may use Short Payment (SP), Allow Deductions (AL), Short Shipped (SS), Bankruptcy (BK), Partial Payment (PT), Cancel (CX) or Unauthorized Deduction (UD). Checked on the memo code (C9) and the memo reason text (A9). The same list (`app/blocked_reasons.py`) is frozen in Create batch. | Choose another reason code | BLOCKED |

## Stages

- Batch format: protocols 12
- Batch values and Ageing: protocols 2, 26, 17, 27, 28, 29, 30, 31, 32
- All other protocols form the last stage.

Notes only (never an error; shown in the Notes list with the error type and the text):
- "Same style, different price": the same style no. with a different price on another memo sheet (a price above the Sale Report is still an error by protocol 10).
- "Column width differs from the template": a column width that is not the template's (protocol 12).
- "Sale Report not uploaded": no Sale Report of the account is uploaded, so protocols 4, 10 and 30 are not checked; upload it and check again.
- Also: spaces in the REASON sheet, Batch number format, style no. not in capitals, Batch COMPANY differs from SOLD TO, empty Batch DATE.
