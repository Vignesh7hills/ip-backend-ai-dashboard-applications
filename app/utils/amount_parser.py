"""
Utilities for parsing Indian-format currency amounts from PDF/Excel ledger files.

Indian number format uses commas after every 2 digits after the first 3:
  e.g.  1,23,456.78  →  123456.78
        12,60,694.00 →  1260694.00

Balance strings carry a direction suffix: "Cr" | "Dr" | "-" (zero)
"""
import re
from typing import Tuple

_AMOUNT_RE = re.compile(r'^([\d,]+\.\d{2})(Cr|Dr)?$')
_ZERO_RE = re.compile(r'^-+$')


def parse_amount(text: str) -> float:
    """
    Convert an Indian-format amount string to float.
    Handles: "12,60,694.00", "12,60,694.00Cr", "87,628.00", "-"
    Returns 0.0 for zero / empty / unrecognised.
    """
    if not text:
        return 0.0
    text = text.strip()
    if _ZERO_RE.match(text):
        return 0.0
    cleaned = text.replace(',', '').rstrip('CrDr').rstrip()
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def parse_balance(text: str) -> Tuple[float, str]:
    """
    Parse a balance string into (amount, direction).

    Handles both concatenated ("12,60,694.00Cr") and
    space-separated ("277000.00 Cr") direction suffixes.

    Returns:
        (0.0, "Zero")    for "-"
        (amount, "Cr")   for "12,60,694.00Cr" or "277000.00 Cr"
        (amount, "Dr")   for "7,25,007.00Dr"  or "277000.00 Dr"
        (amount, "Cr")   if no suffix (treat as Cr by default)
    """
    if not text:
        return 0.0, "Zero"
    text = text.strip()
    if _ZERO_RE.match(text):
        return 0.0, "Zero"
    # Space-separated direction: "277000.00 Cr" / "277000.00 Dr"
    if text.upper().endswith(' CR'):
        return parse_amount(text[:-3]), "Cr"
    if text.upper().endswith(' DR'):
        return parse_amount(text[:-3]), "Dr"
    # Concatenated direction: "12,60,694.00Cr" / "7,25,007.00Dr"
    if text.endswith('Cr'):
        return parse_amount(text[:-2]), "Cr"
    if text.endswith('Dr'):
        return parse_amount(text[:-2]), "Dr"
    # No suffix – could be plain number (treat as Cr)
    val = parse_amount(text)
    return val, ("Zero" if val == 0.0 else "Cr")


def cr_value(amount: float, direction: str) -> float:
    """
    Return the signed Cr value.
    Cr balance → positive; Dr balance → negative; Zero → 0.
    """
    if direction == "Cr":
        return amount
    if direction == "Dr":
        return -amount
    return 0.0


def format_indian(value: float) -> str:
    """Format a float as Indian currency string (no symbol)."""
    if value == 0:
        return "0.00"
    is_negative = value < 0
    value = abs(value)
    integer_part = int(value)
    decimal_part = f"{value:.2f}".split('.')[1]

    s = str(integer_part)
    if len(s) <= 3:
        result = s
    else:
        result = s[-3:]
        s = s[:-3]
        while s:
            result = s[-2:] + ',' + result
            s = s[:-2]
    return ('-' if is_negative else '') + result + '.' + decimal_part
