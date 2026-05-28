"""
Date parsing utilities for various formats encountered in financial files.
"""
import re
from datetime import date, datetime
from typing import Optional

_FORMATS = [
    "%d-%m-%Y",   # 17-10-2024  (ledger PDFs)
    "%d/%m/%Y",   # 17/10/2024
    "%Y-%m-%d",   # 2024-10-17  (ISO)
    "%d-%b-%Y",   # 17-Oct-2024
    "%d %b %Y",   # 17 Oct 2024
    "%d-%m-%y",   # 17-10-24
]

_DATE_LINE_RE = re.compile(r'^\d{2}[-/]\d{2}[-/]\d{4}$')


def is_date_line(text: str) -> bool:
    """Return True if the line looks like a dd-mm-yyyy or dd/mm/yyyy date."""
    return bool(_DATE_LINE_RE.match(text.strip()))


def parse_date(text: str) -> Optional[date]:
    """Try all known formats and return a date object, or None."""
    if not text:
        return None
    text = text.strip()
    for fmt in _FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def fiscal_year_of(d: date, fy_start_month: int = 4) -> str:
    """Return 'FY2024-25' style string for the given date."""
    if d.month >= fy_start_month:
        return f"FY{d.year}-{str(d.year + 1)[-2:]}"
    return f"FY{d.year - 1}-{str(d.year)[-2:]}"
