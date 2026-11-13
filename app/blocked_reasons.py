"""Reason codes that may NOT be used on a credit memo (frozen when creating a batch, an error when rechecking)."""
import re

# (code, description) as in the REASON sheet; a memo is blocked when its code OR its reason text matches.
BLOCKED = [
    ("SP", "Short Payment"),
    ("AL", "Allow Deductions"),
    ("SS", "Short Shipped"),
    ("BK", "Bankruptcy"),
    ("PT", "Partial Payment"),
    ("CX", "Cancel"),
    ("UD", "Unauthorized Deduction"),
]
_CODES = {c for c, _ in BLOCKED}
_TEXTS = {re.sub(r"\s+", " ", d).strip().upper() for _, d in BLOCKED} | {"ALLOW DEDUCTION"}


def is_blocked(code="", description="") -> bool:
    code = (code or "").strip().upper()
    text = re.sub(r"\s+", " ", description or "").strip().upper()
    return (bool(code) and code in _CODES) or (bool(text) and text in _TEXTS)
