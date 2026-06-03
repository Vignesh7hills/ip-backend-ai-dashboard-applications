"""
TDS Returns Parser — handles both Vanshika (TDS Register) and
Shridhar Mills (Summary Register for TDS) XLS formats.

Vanshika format (10-11 cols):
  R12: SR | PARTY NAME | ... | PAN NO. | ... | TDS % | CR.  AMOUNT | TDS | CHALLAN NO.
  R13: T. D. S. Payable A/c. ( Contract )   ← section row (after header)
  R15: FOR PERIOD  2025-10-01  TO  2025-10-31
  R17+: data rows (every other row — sparse layout)

Shridhar format (27 cols):
  R10: PARTY NAME | TYPE | % | PAN No | BILL AMOUNT | CR. AMOUNT | TDS AMT
  R12: TDS A/C : TDS PAYABLE ( 194 C )       ← section row
  R14+: data rows
"""

import re
import calendar
from datetime import date
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass, field
import pandas as pd
from app.core.logger import get_logger
from app.core.exceptions import FileParseError, EmptyFileError, UnsupportedFileTypeError
from app.utils.amount_parser import parse_amount

logger = get_logger(__name__)

_PAN_RE = re.compile(r'^[A-Z]{5}\d{4}[A-Z]$')


def _last_day(year: int, month: int) -> str:
    day = calendar.monthrange(year, month)[1]
    return f"{day:02d}/{month:02d}/{year}"


def _parse_ym(text: str) -> Optional[Tuple[int, int]]:
    """Extract (year, month) from date strings."""
    # 2025-12-01 format
    m = re.search(r'(\d{4})-(\d{2})-\d{2}', text)
    if m:
        return int(m.group(1)), int(m.group(2))
    # 01/10/2025 format
    m = re.search(r'(\d{2})/(\d{2})/(\d{4})', text)
    if m:
        return int(m.group(3)), int(m.group(2))
    return None


def _detect_section(text: str) -> str:
    """Extract TDS section from text."""
    t = text.upper()
    # Exact section codes — 194C, 194Q, 194A, etc.
    m = re.search(r'\b(194\s*[A-Z]?[A-Z]?)\b', t)
    if m:
        return m.group(1).replace(' ', '')
    # '94 Q' or '94Q'
    m = re.search(r'\b94\s*([A-Z])\b', t)
    if m:
        return '194' + m.group(1)
    # '194' alone
    if re.search(r'\b194\b', t):
        # Guess from keywords
        if 'INTEREST' in t:
            return '194A'
        if 'CONTRACT' in t:
            return '194C'
        if 'PURCHASE' in t:
            return '194Q'
        return '194C'  # default
    # Keyword fallback
    if 'INTEREST' in t:
        return '194A'
    if 'CONTRACT' in t:
        return '194C'
    if 'PROFESSIONAL' in t:
        return '194J'
    return 'UNKNOWN'


@dataclass
class TDSEntry:
    deductee_name: str
    pan: str
    section: str
    amount_paid: float
    tds_deducted: float
    rate: float
    payment_date_str: str
    challan_serial: int = 0
    tds_deposited: float = 0.0
    pan_valid: bool = False
    expected_tds: float = 0.0
    tds_shortfall: float = 0.0
    source_file: str = ''
    source_group: str = ''


def _sv(row: list, col: int) -> str:
    v = str(row[col]).strip() if 0 <= col < len(row) else ''
    return '' if v.lower() == 'nan' else v


class TDSReturnsParser:

    def parse_file(self, file_path: str) -> List[TDSEntry]:
        ext = file_path.rsplit('.', 1)[-1].lower()
        try:
            if ext in ('xlsx', 'xls'):
                engine = 'xlrd' if ext == 'xls' else 'openpyxl'
                df = pd.read_excel(file_path, header=None, dtype=str, engine=engine)
            elif ext == 'csv':
                df = pd.read_csv(file_path, header=None, dtype=str)
            elif ext == 'pdf':
                df = self._pdf_to_dataframe(file_path)
            else:
                raise UnsupportedFileTypeError(f"Unsupported: {ext}")
        except (UnsupportedFileTypeError, FileParseError):
            raise
        except Exception as e:
            raise FileParseError(f"Failed to read TDS file: {e}")

        df = df.fillna('').map(lambda x: str(x).strip())
        fname = file_path.rsplit('/', 1)[-1]

        entries = self._parse_shridhar(df, fname)
        if entries:
            logger.info("Parsed %d entries (Shridhar format) from %s", len(entries), fname)
            return entries

        entries = self._parse_vanshika(df, fname)
        if entries:
            logger.info("Parsed %d entries (Vanshika format) from %s", len(entries), fname)
            return entries

        entries = self._parse_jai_kanhaiya(df, fname)
        if entries:
            logger.info("Parsed %d entries (Jai Kanhaiya format) from %s", len(entries), fname)
            return entries

        raise FileParseError(
            "Could not detect TDS columns. Expected:\n"
            "  PARTY NAME | % | PAN No | CR. AMOUNT | TDS AMT  (Shridhar)\n"
            "  SR | PARTY NAME | PAN NO. | TDS % | CR. AMOUNT | TDS  (Vanshika)\n"
            "  Sr. | Account Name | PAN | Applicable Amt. | TDS Amt. | Challan no.  (Jai Kanhaiya)"
        )

    # ── Vanshika TDS Register ─────────────────────────────────────────────────
    # Header at row with 'SR', 'PARTY NAME', 'PAN NO.', 'TDS %', 'CR.  AMOUNT', 'TDS'

    def _parse_vanshika(self, df: pd.DataFrame, fname: str) -> List[TDSEntry]:
        header_row = -1
        col_sr = col_name = col_pan = col_rate = col_cr = col_tds = -1

        for ri in range(min(25, len(df))):
            row = [str(c).lower().strip() for c in df.iloc[ri]]
            # Must have 'party name' and 'tds' and at least 8 columns
            if not any('party name' in c for c in row):
                continue
            if not any(c == 'tds' for c in row):
                continue
            # Find columns
            for ci, c in enumerate(row):
                if c in ('sr', 'sr.') and col_sr < 0:
                    col_sr = ci
                if 'party name' in c and col_name < 0:
                    col_name = ci
                if 'pan no' in c and col_pan < 0:
                    col_pan = ci
                if ('tds %' in c or (c.startswith('%') and len(c) < 6)) and col_rate < 0:
                    col_rate = ci
                if 'cr.' in c and col_cr < 0:
                    col_cr = ci
                # 'cr' alone (col header might just be 'cr.')
                if c == 'cr.  amount' and col_cr < 0:
                    col_cr = ci
                if c == 'tds' and col_tds < 0:
                    col_tds = ci

            # Also detect by position if 'cr. amount' not found exactly
            if col_cr < 0:
                for ci, c in enumerate(row):
                    if 'amount' in c and col_cr < 0:
                        col_cr = ci

            if col_name >= 0 and col_cr >= 0 and col_tds >= 0:
                header_row = ri
                break

        if header_row < 0:
            return []

        # Scan header row more carefully using original (not lower-cased)
        raw_header = [str(c).strip() for c in df.iloc[header_row]]
        for ci, c in enumerate(raw_header):
            cl = c.lower()
            if 'sr' == cl or cl == 'sr.' and col_sr < 0:
                col_sr = ci
            if 'party name' in cl and col_name < 0:
                col_name = ci
            if 'pan no' in cl and col_pan < 0:
                col_pan = ci
            if 'tds %' in cl and col_rate < 0:
                col_rate = ci
            if 'cr.' in cl and col_cr < 0:
                col_cr = ci
            if cl == 'tds' and col_tds < 0:
                col_tds = ci

        entries: List[TDSEntry] = []
        current_section = 'UNKNOWN'
        current_year = current_month = None
        current_group = ''

        for ri in range(header_row + 1, len(df)):
            row = list(df.iloc[ri])
            row_str = ' '.join(str(c).strip() for c in row if str(c).strip() not in ('nan', ''))

            # Period row
            if 'FOR PERIOD' in row_str.upper():
                for cell in row:
                    ym = _parse_ym(str(cell))
                    if ym and ym[1] > 0:
                        # Use the end-of-period month (the second date)
                        pass
                # Parse both start and end dates; use end date's month
                dates_found = [_parse_ym(str(c)) for c in row if _parse_ym(str(c))]
                if dates_found:
                    # Take the later month (end of period)
                    current_year, current_month = max(dates_found, key=lambda x: (x[0], x[1]))
                continue

            # Section row
            row_low = row_str.lower()
            if ('t. d. s.' in row_low or 'tds payable' in row_low or
                    'tds a/c' in row_low or 'payable a/c' in row_low):
                current_section = _detect_section(row_str)
                current_group = row_str.strip()
                continue

            # Skip totals and header repeats
            if 'total' in row_low or ('party name' in row_low and 'for period' not in row_low):
                continue

            # Data row: must have a numeric SR in col 0
            sr = _sv(row, col_sr) if col_sr >= 0 else _sv(row, 0)
            name = _sv(row, col_name) if col_name >= 0 else ''

            if not sr.isdigit() or not name:
                continue
            if name.lower() in ('nan', ''):
                continue
            # Skip special entries
            if any(kw in name.lower() for kw in ('late payment interest', 'interest paid a/c',
                                                   'total', 'summary', 'party name')):
                continue

            pan  = _sv(row, col_pan).upper()  if col_pan  >= 0 else ''
            rate = parse_amount(_sv(row, col_rate)) if col_rate >= 0 else 0.0
            cr   = parse_amount(_sv(row, col_cr))   if col_cr   >= 0 else 0.0
            tds  = parse_amount(_sv(row, col_tds))  if col_tds  >= 0 else 0.0

            if cr == 0.0 and tds == 0.0:
                continue

            # Infer section from filename/rate if still UNKNOWN
            section = current_section
            if section == 'UNKNOWN':
                if '94q' in fname.lower() or '94 q' in fname.lower():
                    section = '194Q'
                elif '94c' in fname.lower() or '94 c' in fname.lower():
                    section = '194C'
                elif rate == 10.0:
                    section = '194A'
                elif rate in (1.0, 2.0):
                    section = '194C'
                elif rate == 0.1:
                    section = '194Q'

            pay_date = _last_day(current_year, current_month) if current_year and current_month else ''

            e = TDSEntry(
                deductee_name=name, pan=pan, section=section,
                amount_paid=cr, tds_deducted=tds, rate=rate,
                payment_date_str=pay_date, source_file=fname,
                source_group=f"{section}_{fname}",
            )
            e.pan_valid = bool(_PAN_RE.match(pan)) if pan else False
            entries.append(e)

        return entries

    # ── Shridhar Summary Register ─────────────────────────────────────────────
    # Header at row with PARTY NAME, TYPE, %, PAN No, BILL AMOUNT, CR. AMOUNT, TDS AMT

    def _parse_shridhar(self, df: pd.DataFrame, fname: str) -> List[TDSEntry]:
        header_row = -1
        col_name = col_type = col_rate = col_pan = col_cr = col_tds = -1

        for ri in range(min(25, len(df))):
            row = [str(c).lower().strip() for c in df.iloc[ri]]
            has_name = any('party name' in c for c in row)
            has_tds  = any('tds amt' in c or 'tax amt' in c for c in row)
            has_pan  = any('pan no' in c for c in row)
            if has_name and has_tds and has_pan:
                for ci, c in enumerate(row):
                    if 'party name' in c and col_name < 0: col_name = ci
                    if c == 'type' and col_type < 0: col_type = ci
                    if c == '%' and col_rate < 0: col_rate = ci
                    if 'pan' in c and 'no' in c and col_pan < 0: col_pan = ci
                    if 'cr.' in c and 'amount' in c and col_cr < 0: col_cr = ci
                    if ('tds amt' in c or 'tax amt' in c) and col_tds < 0: col_tds = ci
                if col_name >= 0 and col_tds >= 0:
                    header_row = ri
                    break

        if header_row < 0:
            return []

        entries: List[TDSEntry] = []
        current_section = 'UNKNOWN'
        current_year = current_month = None
        current_group_label = ''

        # Pre-scan full file for the period date (may appear before header row)
        for ri in range(len(df)):
            row = list(df.iloc[ri])
            row_str = ' '.join(str(c).strip() for c in row if str(c).strip() not in ('nan', ''))
            if 'from date' in row_str.lower() or 'to date' in row_str.lower():
                m_end = re.search(r'[Tt]o\s+[Dd]ate\s+(\d{2}/\d{2}/\d{4})', row_str)
                if m_end:
                    ym = _parse_ym(m_end.group(1))
                    if ym:
                        current_year, current_month = ym
                        break
                # Try any date cell
                for cell in row:
                    ym = _parse_ym(str(cell))
                    if ym:
                        current_year, current_month = ym
                        break
                if current_year:
                    break

        for ri in range(header_row + 1, len(df)):
            row = list(df.iloc[ri])
            row_str = ' '.join(str(c).strip() for c in row if str(c).strip() not in ('nan', ''))
            row_low = row_str.lower()

            # Period row (in case there are multiple sections with different months)
            if 'from date' in row_low:
                m_end = re.search(r'[Tt]o\s+[Dd]ate\s+(\d{2}/\d{2}/\d{4})', row_str)
                if m_end:
                    ym = _parse_ym(m_end.group(1))
                    if ym: current_year, current_month = ym
                continue

            # Section row
            if 'tds a/c' in row_low or 'tds payable' in row_low:
                current_section = _detect_section(row_str)
                # Build unique group label including any suffix like -2
                m_sfx = re.search(r'\(([^)]+)\)(?:-(\d+))?', row_str)
                sfx = '-' + m_sfx.group(2) if (m_sfx and m_sfx.group(2)) else ''
                current_group_label = f"{current_section}{sfx}"
                continue

            # Skip total/header rows
            if 'total' in row_low or 'party name' in row_low:
                continue

            name = _sv(row, col_name) if col_name >= 0 else ''
            if not name or name.lower() in ('nan', ''):
                continue
            if any(kw in name.lower() for kw in ('total', 'party name', 'summary')):
                continue

            pan  = _sv(row, col_pan).upper() if col_pan  >= 0 else ''
            rate = parse_amount(_sv(row, col_rate))  if col_rate >= 0 else 0.0
            cr   = parse_amount(_sv(row, col_cr))    if col_cr   >= 0 else 0.0
            tds  = parse_amount(_sv(row, col_tds))   if col_tds  >= 0 else 0.0

            if cr == 0.0 and tds == 0.0:
                continue

            section = current_section
            if section == 'UNKNOWN' and rate > 0:
                if rate == 10.0: section = '194A'
                elif rate in (1.0, 2.0): section = '194C'
                elif rate == 0.1: section = '194Q'

            pay_date = _last_day(current_year, current_month) if current_year and current_month else ''

            e = TDSEntry(
                deductee_name=name, pan=pan, section=section,
                amount_paid=cr, tds_deducted=tds, rate=rate,
                payment_date_str=pay_date, source_file=fname,
                source_group=f"{fname}_{current_group_label}",
            )
            e.pan_valid = bool(_PAN_RE.match(pan)) if pan else False
            entries.append(e)

        return entries


    # ── Jai Kanhaiya / New Format ─────────────────────────────────────────────
    # Structure:
    #   R6:  Company name
    #   R10: PAN  :  <pan>
    #   R12: TDS (194Q) Purchase Details   ...   From : DD/MM/YYYY  To DD/MM/YYYY
    #   R14: Sr. | Account Name | PAN | Applicable Amt. | TDS Amt. | Challan no.
    #   R16: 1 | party | pan | amount | tds | challan   (every other row — sparse)
    #   R20: '' | '' | Total | total_amt | total_tds | ''
    #
    # Key differences from Vanshika:
    #   - "Account Name" instead of "Party Name"
    #   - "Applicable Amt." instead of "CR. Amount"
    #   - "TDS Amt." instead of "TDS" (exact)
    #   - Section + date both on the same pre-header row
    #   - No separate rate (%) column

    def _parse_jai_kanhaiya(self, df: pd.DataFrame, fname: str) -> List[TDSEntry]:
        header_row = -1
        col_sr = col_name = col_pan = col_amt = col_tds = col_challan = -1

        for ri in range(min(30, len(df))):
            row = [str(c).lower().strip() for c in df.iloc[ri]]
            has_name   = any('account name' in c or 'party name' in c for c in row)
            has_tds    = any('tds amt' in c or 'tds amount' in c for c in row)
            has_pan    = any(c == 'pan' or 'pan no' in c for c in row)
            has_amount = any('applicable amt' in c or 'applicable amount' in c
                             or 'amount' in c for c in row)
            if has_name and has_tds and has_pan and has_amount:
                for ci, c in enumerate(row):
                    if c in ('sr', 'sr.') and col_sr < 0:            col_sr     = ci
                    if ('account name' in c or 'party name' in c) and col_name < 0:
                        col_name = ci
                    if (c == 'pan' or 'pan no' in c) and col_pan < 0: col_pan    = ci
                    if ('applicable amt' in c or 'applicable amount' in c) and col_amt < 0:
                        col_amt = ci
                    if ('tds amt' in c or 'tds amount' in c) and col_tds < 0:
                        col_tds = ci
                    if 'challan' in c and col_challan < 0:            col_challan = ci
                if col_name >= 0 and col_amt >= 0 and col_tds >= 0:
                    header_row = ri
                    break

        if header_row < 0:
            return []

        # Pre-scan all rows above header for section and date
        current_section = 'UNKNOWN'
        current_year = current_month = None

        for ri in range(header_row):
            row = list(df.iloc[ri])
            row_str = ' '.join(str(c).strip() for c in row
                               if str(c).strip() not in ('nan', '', 'None'))
            if not row_str:
                continue

            # Section detection: "TDS (194Q) Purchase Details"
            if re.search(r'\b194', row_str.upper()):
                sec = _detect_section(row_str)
                if sec != 'UNKNOWN':
                    current_section = sec

            # Date: "From : 01/03/2026  To 31/03/2026"
            if 'from' in row_str.lower() or 'to' in row_str.lower():
                # Collect all dates in the row and pick the latest (end date)
                dates = [_parse_ym(str(c)) for c in row
                         if _parse_ym(str(c)) is not None]
                if not dates:
                    # Also try the full row_str for "DD/MM/YYYY" patterns
                    for m in re.finditer(r'(\d{2}/\d{2}/\d{4})', row_str):
                        ym = _parse_ym(m.group(1))
                        if ym:
                            dates.append(ym)
                if dates:
                    current_year, current_month = max(dates, key=lambda x: (x[0], x[1]))

        entries: List[TDSEntry] = []

        for ri in range(header_row + 1, len(df)):
            row = list(df.iloc[ri])
            row_str = ' '.join(str(c).strip() for c in row
                               if str(c).strip() not in ('nan', '', 'None'))
            row_low = row_str.lower()

            # Skip total rows
            if 'total' in row_low:
                continue

            # Data row: sr column (or col 0) must be a digit
            sr_val = _sv(row, col_sr) if col_sr >= 0 else _sv(row, 0)
            if not sr_val.isdigit():
                continue

            name = _sv(row, col_name) if col_name >= 0 else ''
            if not name or name.lower() in ('nan', ''):
                continue
            if any(kw in name.lower() for kw in ('total', 'summary')):
                continue

            pan = _sv(row, col_pan).upper() if col_pan >= 0 else ''
            amt = parse_amount(_sv(row, col_amt)) if col_amt >= 0 else 0.0
            tds = parse_amount(_sv(row, col_tds)) if col_tds >= 0 else 0.0

            if amt == 0.0 and tds == 0.0:
                continue

            # Infer rate from amount and TDS (tds / amt * 100)
            rate = round(tds / amt * 100, 4) if amt > 0 and tds > 0 else 0.0

            pay_date = (_last_day(current_year, current_month)
                        if current_year and current_month else '')

            section = current_section
            # Fallback from filename
            if section == 'UNKNOWN':
                if '94q' in fname.lower():
                    section = '194Q'
                elif '94c' in fname.lower():
                    section = '194C'
                elif '94a' in fname.lower():
                    section = '194A'

            e = TDSEntry(
                deductee_name=name, pan=pan, section=section,
                amount_paid=amt, tds_deducted=tds, rate=rate,
                payment_date_str=pay_date, source_file=fname,
                source_group=f"{section}_{fname}",
            )
            e.pan_valid = bool(_PAN_RE.match(pan)) if pan else False
            entries.append(e)

        if entries:
            logger.info("Parsed %d entries (Jai Kanhaiya format) from %s",
                        len(entries), fname)
        return entries

    def _pdf_to_dataframe(self, file_path: str) -> pd.DataFrame:
        try:
            from pypdf import PdfReader
        except ImportError:
            raise FileParseError("pypdf not installed.")
        all_rows = []
        reader = PdfReader(file_path)
        for page in reader.pages:
            for line in (page.extract_text() or '').split('\n'):
                line = line.strip()
                if line:
                    all_rows.append(re.split(r'\s{2,}', line))
        if not all_rows:
            raise EmptyFileError("No text found in PDF.")
        max_cols = max(len(r) for r in all_rows)
        padded = [r + [''] * (max_cols - len(r)) for r in all_rows]
        return pd.DataFrame(padded, dtype=str)