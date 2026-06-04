"""
TDS Returns Parser — AUTO-ADAPTIVE column detection.

Instead of hard-coding per-format parsers, the engine:
  1. Scans every row in the first 40 rows for a "header row" by trying to
     match columns to semantic buckets (NAME, PAN, AMOUNT, TDS, RATE, TYPE, DATE).
  2. Once a header row with at least NAME + (AMOUNT or RATE) + (TDS or RATE) is
     found, all subsequent rows are treated as data rows.
  3. Missing values (e.g. no TDS-amount column → compute from amount × rate,
     no section column → infer from TYPE cell or pre-header keywords) are derived
     automatically.

Handles every known format without any format-specific code path:

  Format A  (Vanshika TDS Register)
    SR | PARTY NAME | ... | PAN NO. | TDS % | CR.  AMOUNT | TDS | CHALLAN NO.

  Format B  (Shridhar Summary Register)
    PARTY NAME | TYPE | % | PAN No | BILL AMOUNT | CR. AMOUNT | TDS AMT

  Format C  (Jai Kanhaiya)
    Sr. | Account Name | PAN | Applicable Amt. | TDS Amt. | Challan no.
    Section + date appear in a pre-header row.

  Format D  (Flat / Generic — new format)
    Party | PAN | Type | Amount | TDS %
    No header metadata. TDS amount derived as Amount × Rate / 100.
    Section inferred from "Type" cell value.

  Any future variation with different column names is handled by extending
  the alias lists at the top of the file — no new parser class needed.
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

# ── Column alias tables ────────────────────────────────────────────────────────
# Each list is ordered MOST-SPECIFIC → LEAST-SPECIFIC so the matcher always
# prefers unambiguous terms.

_NAME_ALIASES = [
    'party name', 'account name', 'deductee name', 'vendor name',
    'name of deductee', 'party', 'name', 'deductee', 'vendor',
]

_PAN_ALIASES = [
    'permanent account number', 'pan no.', 'pan no', 'pan number', 'pan',
]

_AMT_ALIASES = [
    'cr.  amount', 'cr. amount', 'applicable amount', 'applicable amt.',
    'applicable amt', 'bill amount', 'taxable amount', 'payment amount',
    'amount paid', 'amount of payment', 'cr.', 'amount',
]

# TDS-amount column: must contain 'tds' (to avoid grabbing the plain 'amount' col)
_TDS_AMT_ALIASES = [
    'tds amount', 'tds amt.', 'tds amt', 'tax deducted', 'tax amount',
    'amount of tds', 'tds deducted', 'tax amt', 'tds',
]

# Rate column: 'tds %' / '%' / 'rate' — must NOT match 'tds amt'
_RATE_ALIASES = [
    'tds %', 'tds rate %', 'tds rate', 'rate of tds', 'rate %', 'rate', '%',
]

# Section/type column: cell value is parsed for section code
_TYPE_ALIASES = [
    'nature of payment', 'tds section', 'section code', 'section',
    'nature', 'type of payment', 'payment type', 'type',
]

# Date column: used in flat/generic format
_DATE_ALIASES = [
    'payment date', 'date of payment', 'period', 'month', 'date',
]

# Challan column (optional)
_CHALLAN_ALIASES = ['challan no.', 'challan no', 'challan number', 'challan']

# SR column
# SR aliases — exact match only (substring match would catch 'PAN No', 'Challan No' etc.)
_SR_ALIASES = ['sr.', 'sr', 's.no.', 's.no', 'sno', 's. no.', 's. no', '#']

# ── Type-value → TDS section mapping ──────────────────────────────────────────
_TYPE_SECTION: Dict[str, str] = {
    # Professional / Technical Fees → 94J
    'professional fees':        '94J',
    'professional':             '94J',
    'technical fees':           '94J',
    'technical service':        '94J',

    # Contractor / Sub-contractor → 94C
    'contract':                 '94C',
    'contractor':               '94C',
    'sub-contractor':           '94C',
    'sub contractor':           '94C',
    'works contract':           '94C',
    'labour contract':          '94C',
    'job work':                 '94C',

    # Purchase of Goods → 94Q
    'purchase of goods':        '94Q',
    'purchase':                 '94Q',
    'goods':                    '94Q',
    'tds on purchase':          '94Q',

    # Interest (other than securities) → 94A
    'interest':                 '94A',
    'interest on loan':         '94A',
    'interest on fd':           '94A',
    'interest on deposit':      '94A',

    # Commission / Brokerage → 94H
    'commission':               '94H',
    'brokerage':                '94H',
    'commission and brokerage': '94H',

    # Rent — Plant/Machinery → 194I(a)
    'rent of plant':            '194I(a)',
    'rent plant':               '194I(a)',
    'rent machinery':           '194I(a)',
    'rent of machinery':        '194I(a)',
    'rent of equipment':        '194I(a)',
    'plant rent':               '194I(a)',
    'machinery rent':           '194I(a)',

    # Rent — Land/Building → 194I(b)
    'rent of land':             '194I(b)',
    'rent of building':         '194I(b)',
    'rent land':                '194I(b)',
    'rent building':            '194I(b)',
    'office rent':              '194I(b)',
    'shop rent':                '194I(b)',
    'house rent':               '194I(b)',
    'rental':                   '194I(b)',
    'rent':                     '194I(b)',

    # Additional real-world variants
    'salary':                   '192',
    'salaries':                 '192',
    'dividend':                 '194',
    'lottery':                  '194B',
    'winnings':                 '194B',
    'transport':                '94C',     # transporter treated as contractor
    'freight':                  '94C',
    'clearing and forwarding':  '94C',
    'c&f':                      '94C',
    'insurance commission':     '94H',
    'advertising':              '94C',
}

# ── Rate → section fallback ────────────────────────────────────────────────────
def _rate_to_section(rate: float) -> str:
    """
    Infer TDS section code from rate — used only when no type/section column exists.
    Rates that map to a single section unambiguously are returned directly.
    Ambiguous rates (shared by multiple sections) return '[REVIEW SECTION]'.

    Standard TDS rates (Income Tax Act 1961):
      0.1% → 94Q   Purchase of Goods           (unique)
      1.0% → 94C   Contractor individual/HUF   (unique)
      2.0% → ambiguous: 94C (contractor company) OR 194I(a) (rent plant)
      5.0% → 94H   Commission/Brokerage        (unique)
      10%  → ambiguous: 94J (professional) OR 94A (interest) OR 194I(b) (rent land)
    """
    if rate == 0.1:  return '94Q'              # Purchase of Goods — unique
    if rate == 1.0:  return '94C'              # Contractor individual/HUF — unique
    if rate == 5.0:  return '94H'              # Commission/Brokerage — unique
    # Ambiguous rates — cannot determine section from rate alone
    if rate == 2.0:  return '[REVIEW SECTION]' # 94C or 194I(a)
    if rate == 10.0: return '[REVIEW SECTION]' # 94J or 94A or 194I(b)
    return 'UNKNOWN'


def _last_day(year: int, month: int) -> str:
    day = calendar.monthrange(year, month)[1]
    return f"{day:02d}/{month:02d}/{year}"


def _parse_ym(text: str) -> Optional[Tuple[int, int]]:
    m = re.search(r'(\d{4})-(\d{2})-\d{2}', text)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.search(r'(\d{2})/(\d{2})/(\d{4})', text)
    if m:
        return int(m.group(3)), int(m.group(2))
    return None


def _detect_section(text: str) -> str:
    """
    Extract TDS section code from a text string (header row, inline section label etc).
    Returns the exact code as used in Form 26Q: 94C, 94A, 94H, 94J, 94Q, 194I(a), 194I(b).
    """
    t = text.upper()

    # Explicit 194I(a) / 194I(b) — check before generic 194I
    if re.search(r'194\s*I\s*\(\s*A\s*\)', t): return '194I(a)'
    if re.search(r'194\s*I\s*\(\s*B\s*\)', t): return '194I(b)'
    if re.search(r'194\s*I\b', t):
        # Determine sub-type from context
        if any(k in t for k in ('PLANT', 'MACHINERY', 'EQUIPMENT')):
            return '194I(a)'
        return '194I(b)'  # default to land/building

    # Explicit 94x codes in the text (e.g. "TDS (94Q) Purchase Details")
    m = re.search(r'\b(94\s*[A-Z])\b', t)
    if m:
        return m.group(1).replace(' ', '')   # e.g. "94C", "94Q"

    # Full 194x codes in the text
    m = re.search(r'\b(194\s*[A-Z])\b', t)
    if m:
        code = m.group(1).replace(' ', '')   # e.g. "194C" → map to "94C"
        mapping = {'194A': '94A', '194C': '94C', '194H': '94H',
                   '194J': '94J', '194Q': '94Q'}
        return mapping.get(code, code)

    # Keyword-based fallback — ordered most-specific to least-specific
    if 'PROFESSIONAL' in t or 'TECHNICAL SERVICE' in t: return '94J'
    if 'INTEREST'     in t:                              return '94A'
    if 'COMMISSION'   in t or 'BROKERAGE' in t:         return '94H'
    if 'PURCHASE'     in t or 'GOODS'     in t:         return '94Q'
    if 'CONTRACT'     in t or 'CONTRACTOR' in t:        return '94C'
    if any(k in t for k in ('RENT', 'RENTAL')):
        if any(k in t for k in ('PLANT', 'MACHINERY', 'EQUIPMENT')): return '194I(a)'
        return '194I(b)'
    if 'SALARY' in t or 'SALARIES' in t:                return '192'
    if 'DIVIDEND' in t:                                  return '194'
    if 'LOTTERY' in t or 'WINNINGS' in t:               return '194B'
    return 'UNKNOWN'


def _type_value_to_section(type_val: str) -> str:
    """
    Infer TDS section from a 'Type / Nature of Payment' cell value.
    Returns '[REVIEW SECTION]' when the type is present but ambiguous,
    or 'UNKNOWN' when no match is found at all.
    """
    t = type_val.lower().strip()
    if not t:
        return 'UNKNOWN'

    # 1. Try direct code detection (e.g. cell contains '94J', '194I(a)')
    sec = _detect_section(type_val)
    if sec not in ('UNKNOWN', '[REVIEW SECTION]'):
        return sec

    # 2. Try keyword table (most-specific match wins)
    #    Iterate in order so longer/more-specific keys match before shorter ones
    for keyword in sorted(_TYPE_SECTION.keys(), key=len, reverse=True):
        if keyword in t:
            return _TYPE_SECTION[keyword]

    # 3. Nothing matched — flag for manual review
    return '[REVIEW SECTION]'


def _sv(row: list, col: int) -> str:
    v = str(row[col]).strip() if 0 <= col < len(row) else ''
    return '' if v.lower() in ('nan', 'none') else v


# ── Column finder ──────────────────────────────────────────────────────────────

def _find_col(header_lower: List[str], aliases: List[str]) -> int:
    """
    Return the first column index that matches any alias.
    Aliases are tried most-specific first.

    Two-pass matching prevents weaker substring matches from stealing
    columns that belong to more-specific aliases.
    E.g. 'tds amount' must not match col 'amount' before it finds col 'tds amount'.

    Pass 1 (strong): alias == h  OR  alias in h  (alias is substring of header)
    Pass 2 (weak):   h in alias                   (header is substring of alias)
    """
    # Pass 1: exact match or alias contained in header cell
    for alias in aliases:
        for ci, h in enumerate(header_lower):
            if not h:
                continue
            if alias == h or alias in h:
                return ci
    # Pass 2: header cell contained in alias (weak — only if nothing found above)
    for alias in aliases:
        for ci, h in enumerate(header_lower):
            if not h:
                continue
            if h in alias:
                return ci
    return -1



# ── Header-row detector ────────────────────────────────────────────────────────

def _detect_header_row(df: pd.DataFrame) -> Tuple[int, Dict[str, int]]:
    """
    Scan rows 0–39 for a row that yields at least:
      • col_name  (party / account / deductee)
      • col_amt OR col_rate
      • col_tds  OR col_rate   (we can compute TDS from amt×rate)

    Returns (header_row_index, col_map) where col_map keys are:
      'name', 'pan', 'amt', 'tds', 'rate', 'type', 'date', 'sr', 'challan'
    Values are column indices (-1 = not found).
    """
    for ri in range(min(40, len(df))):
        row_raw  = [str(c).strip()        for c in df.iloc[ri]]
        row_low  = [c.lower()             for c in row_raw]

        col_name    = _find_col(row_low, _NAME_ALIASES)
        col_pan     = _find_col(row_low, _PAN_ALIASES)
        col_tds     = _find_col(row_low, _TDS_AMT_ALIASES)
        col_rate    = _find_col(row_low, _RATE_ALIASES)
        col_amt     = _find_col(row_low, _AMT_ALIASES)
        col_type    = _find_col(row_low, _TYPE_ALIASES)
        col_date    = _find_col(row_low, _DATE_ALIASES)
        # SR column: exact match only — 'no' must not match 'pan no' or 'challan no'
        col_sr = next(
            (ci for ci, h in enumerate(row_low)
             if h in {'sr.', 'sr', 's.no.', 's.no', 'sno', 's. no.', 's. no', '#'}),
            -1,
        )
        col_challan = _find_col(row_low, _CHALLAN_ALIASES)

        # Collision resolution:
        # 1. col_amt == col_tds: 'tds' alias matched an amount column (e.g. "TDS %")
        #    → recalculate col_tds requiring 'tds' explicitly and different from amt col.
        if col_amt >= 0 and col_tds == col_amt:
            col_tds = next(
                (ci for ci, h in enumerate(row_low)
                 if 'tds' in h and ci != col_amt),
                -1,
            )
        # 2. col_tds == col_rate: the same column matched both aliases.
        #    "TDS %" is a RATE column, not a TDS-amount column.
        #    Reset col_tds so TDS amount gets computed as amt * rate / 100 later.
        if col_tds >= 0 and col_rate >= 0 and col_tds == col_rate:
            col_tds = -1

        # 3. Validate col_amt against first data row.
        #    If the value in that column looks like a date (not a number),
        #    it is actually the date column — reassign and re-scan for the
        #    real amount column in the remaining numeric columns.
        first_data_row = ri + 1
        if col_amt >= 0 and first_data_row < len(df):
            sample = str(df.iloc[first_data_row, col_amt]).strip()
            if re.search(r'\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}|\d{2}\.\d{2}\.\d{4}', sample):
                # col_amt is really the date column
                if col_date < 0:
                    col_date = col_amt
                # Re-scan for amount: first numeric column that isn't col_date/col_pan/col_name
                # Note: col_rate is NOT in reserved — some files put amount under "TDS Rate %" header
                reserved = {col_name, col_pan, col_date, col_tds, col_sr, col_challan}
                col_amt = -1
                for ci2 in range(len(row_raw)):
                    if ci2 in reserved:
                        continue
                    sample2 = str(df.iloc[first_data_row, ci2]).strip()
                    try:
                        v = float(sample2.replace(',', ''))
                        if v > 100:   # amounts are typically > 100
                            col_amt = ci2
                            break
                    except ValueError:
                        pass
                # If col_amt grabbed what was col_rate, clear col_rate to avoid confusion
                if col_amt >= 0 and col_amt == col_rate:
                    col_rate = -1

        # Minimum required: name + (amt or rate) + (tds or rate)
        has_name   = col_name >= 0
        has_value  = col_amt  >= 0 or col_rate >= 0
        has_tds    = col_tds  >= 0 or col_rate >= 0

        if has_name and has_value and has_tds:
            col_map = {
                'name': col_name, 'pan': col_pan, 'amt': col_amt,
                'tds': col_tds, 'rate': col_rate, 'type': col_type,
                'date': col_date, 'sr': col_sr, 'challan': col_challan,
            }
            logger.debug(
                "Header detected at row %d: %s",
                ri, {k: v for k, v in col_map.items() if v >= 0}
            )
            return ri, col_map

    return -1, {}


# ── Pre-header metadata scanner ───────────────────────────────────────────────

def _scan_pre_header(df: pd.DataFrame, header_row: int,
                     fname: str) -> Tuple[str, Optional[int], Optional[int]]:
    """
    Scan rows before the header for:
      - TDS section code (e.g. "TDS (194Q) Purchase Details")
      - Period end date  (e.g. "From : 01/03/2026  To 31/03/2026")
    Returns (section, year, month).
    """
    section = 'UNKNOWN'
    year = month = None

    for ri in range(header_row):
        row_str = ' '.join(
            str(c).strip() for c in df.iloc[ri]
            if str(c).strip() not in ('nan', '', 'None')
        )
        if not row_str:
            continue

        # Section
        if re.search(r'\b194', row_str.upper()):
            sec = _detect_section(row_str)
            if sec != 'UNKNOWN':
                section = sec

        # Period dates — pick the latest
        all_dates = [_parse_ym(str(c)) for c in df.iloc[ri]]
        all_dates += [_parse_ym(m.group()) for m in re.finditer(r'\d{2}/\d{2}/\d{4}', row_str)]
        valid = [d for d in all_dates if d]
        if valid:
            year, month = max(valid, key=lambda x: (x[0], x[1]))

    # Fallback: infer section from filename
    if section == 'UNKNOWN':
        fl = fname.lower()
        if '94q' in fl:    section = '94Q'
        elif '94c' in fl:  section = '94C'
        elif '94a' in fl:  section = '94A'
        elif '94j' in fl:  section = '94J'
        elif '94h' in fl:  section = '94H'

    return section, year, month


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
    bank_challan_no: str = ''   # actual challan no. from source (e.g. 23, 234)



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

        entries = self._parse_adaptive(df, fname)
        if entries:
            return entries

        raise FileParseError(
            f"[{fname}] Could not detect TDS columns.\n"
            "The file must have a header row containing at minimum:\n"
            "  • A party/name column  (Party, Party Name, Account Name, Deductee, etc.)\n"
            "  • An amount column     (Amount, CR. Amount, Applicable Amt, etc.)\n"
            "  • A TDS or rate column (TDS, TDS Amt, TDS %, Rate, etc.)\n"
            "Column names are matched flexibly — exact spelling is not required."
        )

    def _parse_adaptive(self, df: pd.DataFrame, fname: str) -> List[TDSEntry]:
        """
        Single adaptive parser that works for ALL TDS file formats by detecting
        columns semantically rather than matching hard-coded format signatures.
        """
        header_row, cols = _detect_header_row(df)
        if header_row < 0:
            return []

        col_name    = cols['name']
        col_pan     = cols['pan']
        col_amt     = cols['amt']
        col_tds     = cols['tds']
        col_rate    = cols['rate']
        col_type    = cols['type']
        col_date    = cols['date']
        col_sr      = cols['sr']
        col_challan = cols['challan']

        # Pre-header metadata (section, period)
        pre_section, pre_year, pre_month = _scan_pre_header(df, header_row, fname)
        current_section = pre_section
        current_year    = pre_year
        current_month   = pre_month
        current_group   = pre_section  # may be overridden by inline section rows

        entries: List[TDSEntry] = []

        for ri in range(header_row + 1, len(df)):
            row = list(df.iloc[ri])
            row_str = ' '.join(
                str(c).strip() for c in row
                if str(c).strip() not in ('nan', '', 'None')
            )
            row_low = row_str.lower()

            if not row_str:
                continue

            # ── Inline section row (e.g. "TDS A/C : TDS PAYABLE (194C)") ──
            if any(kw in row_low for kw in ('tds a/c', 'tds payable', 'tds payable a/c',
                                             't. d. s.', 'payable a/c')):
                sec = _detect_section(row_str)
                if sec != 'UNKNOWN':
                    current_section = sec
                    current_group   = row_str.strip()
                continue

            # ── Inline period row ──────────────────────────────────────────
            if 'for period' in row_low or 'from date' in row_low:
                all_dates = [_parse_ym(str(c)) for c in row]
                all_dates += [
                    _parse_ym(m.group())
                    for m in re.finditer(r'\d{2}/\d{2}/\d{4}', row_str)
                ]
                valid = [d for d in all_dates if d]
                if valid:
                    current_year, current_month = max(valid, key=lambda x: (x[0], x[1]))
                continue

            # ── Skip total / header-repeat rows ───────────────────────────
            if 'total' in row_low:
                continue
            # Skip rows that look like header repeats (name col = header text)
            name_cell = _sv(row, col_name) if col_name >= 0 else ''
            if name_cell.lower() in [a for a in _NAME_ALIASES]:
                continue

            # ── SR check (optional — skip if SR column found but not numeric) ──
            if col_sr >= 0:
                sr_val = _sv(row, col_sr)
                if sr_val and not sr_val.replace('.', '').isdigit():
                    continue

            # ── Extract values ─────────────────────────────────────────────
            name = name_cell
            if not name or name.lower() in ('nan', ''):
                continue
            if any(kw in name.lower() for kw in ('total', 'summary', 'grand total',
                                                   'sub total', 'party name',
                                                   'account name', 'deductee')):
                continue

            pan  = _sv(row, col_pan).upper()          if col_pan  >= 0 else ''
            amt  = parse_amount(_sv(row, col_amt))     if col_amt  >= 0 else 0.0
            tds  = parse_amount(_sv(row, col_tds))     if col_tds  >= 0 else 0.0
            rate = parse_amount(_sv(row, col_rate))    if col_rate >= 0 else 0.0

            # ── Derive missing TDS amount from amount × rate ───────────────
            if tds == 0.0 and amt > 0 and rate > 0:
                tds = round(amt * rate / 100, 2)

            # ── Derive missing rate from tds / amount ─────────────────────
            if rate == 0.0 and amt > 0 and tds > 0:
                rate = round(tds / amt * 100, 4)

            if amt == 0.0 and tds == 0.0:
                continue

            # ── Section resolution ─────────────────────────────────────────
            section = current_section

            # 1. From type/nature column cell value
            if section == 'UNKNOWN' and col_type >= 0:
                type_val = _sv(row, col_type)
                if type_val:
                    section = _type_value_to_section(type_val)

            # 2. From inferred rate
            if section == 'UNKNOWN' and rate > 0:
                section = _rate_to_section(rate)

            # ── Date resolution ────────────────────────────────────────────
            pay_date = ''
            if col_date >= 0:
                raw_date = _sv(row, col_date)
                ym = _parse_ym(raw_date)
                if ym:
                    pay_date = _last_day(ym[0], ym[1])
            if not pay_date and current_year and current_month:
                pay_date = _last_day(current_year, current_month)

            # ── Challan number from source file ────────────────────────────
            bank_challan = _sv(row, col_challan) if col_challan >= 0 else ''
            if bank_challan and bank_challan.replace('.', '').isdigit():
                bank_challan = str(int(float(bank_challan)))

            # ── source_group: include challan no. so each unique challan
            #    gets its own Challan Serial No. in Form 26Q output ──────
            if bank_challan:
                grp_key = f"{section}_{bank_challan}_{fname}"
            else:
                grp_key = f"{section}_{fname}"

            # ── Build entry ────────────────────────────────────────────────
            e = TDSEntry(
                deductee_name=name,
                pan=pan,
                section=section,
                amount_paid=amt,
                tds_deducted=tds,
                rate=rate,
                payment_date_str=pay_date,
                source_file=fname,
                source_group=grp_key,
                bank_challan_no=bank_challan,
            )
            e.pan_valid      = bool(_PAN_RE.match(pan)) if pan else False
            e.tds_deposited  = tds
            e.expected_tds   = round(amt * rate / 100, 2) if rate else 0.0
            e.tds_shortfall  = max(0.0, e.expected_tds - tds)
            entries.append(e)

        logger.info("Auto-adaptive parser: %d entries from %s", len(entries), fname)
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