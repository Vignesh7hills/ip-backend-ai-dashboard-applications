"""
Utilities for parsing Indian-format currency amounts from PDF/Excel ledger files.

Indian number format uses commas after every 2 digits after the first 3:
  e.g.  1,23,456.78  →  123456.78
        12,60,694.00 →  1260694.00

Balance strings carry a direction suffix: "Cr" | "Dr" | "-" (zero)

Also handles:
  - Parenthetical negatives: (12,345.00) → -12345.00
  - Missing decimal: 12345 → 12345.00
  - Currency symbols: ₹12,345.00, Rs. 12,345
  - Space-separated Cr/Dr: "277000.00 Cr"
  - Excel numeric values (float/int passed as str)
"""
import re
from typing import Tuple

_ZERO_RE   = re.compile(r'^-+$|^0+\.?0*$')
_PAREN_RE  = re.compile(r'^\(([0-9,. ]+)\)$')   # (12,345.00) → negative
_AMOUNT_RE = re.compile(r'^([\\d,]+\\.\\d{2})(Cr|Dr)?$')


def parse_amount(text: str) -> float:
    """
    Convert an Indian-format amount string to float.
    Handles: "12,60,694.00", "12,60,694.00Cr", "87,628.00", "-",
             "(12,345.00)" (negative), "₹1,234", "Rs. 1,234.50",
             plain integers "12345", floats "1234.5"
    Returns 0.0 for zero / empty / unrecognised.
    """
    if text is None:
        return 0.0
    if not isinstance(text, str):
        try:
            return float(text)
        except (ValueError, TypeError):
            return 0.0

    text = text.strip()
    if not text:
        return 0.0
    if _ZERO_RE.match(text):
        return 0.0

    # Parenthetical negative: (12,345.00)
    pm = _PAREN_RE.match(text)
    if pm:
        inner = pm.group(1).replace(',', '').strip()
        try:
            return -float(inner)
        except ValueError:
            return 0.0

    # Strip currency symbols and direction suffixes
    cleaned = text
    cleaned = re.sub(r'[₹$€£]', '', cleaned)     # currency symbols
    cleaned = re.sub(r'\bRs\.?\b', '', cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.strip()
    cleaned = re.sub(r'\s*(Cr|Dr)\s*$', '', cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.replace(',', '').strip()

    # Handle negative prefix
    negative = cleaned.startswith('-')
    if negative:
        cleaned = cleaned.lstrip('-').strip()

    try:
        val = float(cleaned)
        return -val if negative else val
    except ValueError:
        return 0.0


def parse_balance(text: str) -> Tuple[float, str]:
    """
    Parse a balance string into (amount, direction).

    Handles both concatenated ("12,60,694.00Cr") and
    space-separated ("277000.00 Cr") direction suffixes.
    Also handles plain numbers and parenthetical negatives.

    Returns:
        (0.0, "Zero")    for "-" or 0
        (amount, "Cr")   for credit balances
        (amount, "Dr")   for debit balances
        (amount, "Cr")   if no suffix (treat as Cr by default)
    """
    if not text:
        return 0.0, "Zero"
    text = text.strip()
    if not text or _ZERO_RE.match(text):
        return 0.0, "Zero"

    # Space-separated direction: "277000.00 Cr" / "277000.00 Dr"
    upper = text.upper()
    if upper.endswith(' CR'):
        return parse_amount(text[:-3]), "Cr"
    if upper.endswith(' DR'):
        return parse_amount(text[:-3]), "Dr"
    # Concatenated direction: "12,60,694.00Cr" / "7,25,007.00Dr"
    if text.endswith('Cr') or text.upper().endswith('CR'):
        return parse_amount(text[:-2]), "Cr"
    if text.endswith('Dr') or text.upper().endswith('DR'):
        return parse_amount(text[:-2]), "Dr"

    # No suffix — treat as Cr (common in many formats)
    val = parse_amount(text)
    return val, ("Zero" if val == 0.0 else "Cr")


def cr_value(amount: float, direction: str) -> float:
    """
    Return the signed Cr value.
    Cr balance → positive; Dr balance → negative; Zero → 0.
    """
    d = direction.strip().capitalize() if direction else ''
    if d == "Cr":
        return amount
    if d == "Dr":
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
