"""
Trial Balance Parser — FIXED VERSION v2.

Supports four input formats:

  Format A — Standard Dr/Cr TB (Master Reference / any standard TB export):
      Sr No | Ledger Name | Group | Amount (Dr) | Amount (Cr)
      Headers contain 'Amount (Dr)' / 'Amount (Cr)' or 'Debit'/'Credit' or 'Dr'/'Cr'

  Format B — Two-sided BS/P&L with Liabilities/Assets header (UGARARAM files):
      Row with 'Liabilities' + 'Assets' header
      Left side → Credit (Liabilities), Right side → Debit (Assets)

  Format C — Two-sided BS/P&L with PARTICULARS/AMOUNT header (UMA TEXCOM Type 1):
      Row with 'PARTICULARS' repeated twice (left and right side)
      10-column structure: [grp_l, name_l, ?, sub_l, total_l, grp_r, name_r, ?, sub_r, total_r]
      Left side (col 0-4) → Credit (Liabilities side)
      Right side (col 5-9) → Debit (Assets side)

  Format D — PDF two-column layout (UMA TEXCOM bs.pdf / pl.pdf):
      pdfplumber extracts as 2-column table: left_text | right_text
      Each cell contains "NAME  AMOUNT" packed on one line
      Parse each cell as "name  amount" pairs
"""

import re
from typing import List, Tuple, Optional, Dict
import pandas as pd
from app.core.logger import get_logger
from app.core.exceptions import FileParseError, EmptyFileError, UnsupportedFileTypeError
from app.utils.amount_parser import parse_amount

logger = get_logger(__name__)

# ── Keyword lists ─────────────────────────────────────────────────────────────

_DEBIT_KEYWORDS = [
    'amount (dr)', 'amount(dr)', 'dr amount', 'debit amount',
    'debit balance', 'debit (rs.)', 'opening dr', 'closing dr',
    'dr bal', 'debit', 'amount dr', 'dr.',
    'dr',
]
_CREDIT_KEYWORDS = [
    'amount (cr)', 'amount(cr)', 'cr amount', 'credit amount',
    'credit balance', 'credit (rs.)', 'opening cr', 'closing cr',
    'cr bal', 'credit', 'amount cr', 'cr.',
    'cr',
]
_ACCOUNT_KEYWORDS = [
    'ledger name', 'account name', 'account head', 'particulars',
    'description', 'narration', 'ledger', 'account', 'head', 'name',
]
_LIAB_KW = ['liabilit', 'liabilities', 'liablities']
_ASSET_KW = ['asset', 'assets']

_SKIP_ROW_KW = {
    'total', 'grand total', 'sub total', 'subtotal', 'net total',
    'gross profit', 'net profit', 'net loss',
}

_DOTS_RE = re.compile(r'(\s*\.\s*){2,}')
# Pattern to extract trailing amount from "ACCOUNT NAME 123456.78"
# Amount is always at end of line: optional paren/minus, digits, dot, 2 decimal places
# Works with both single-space and multi-space separators (PDF uses single space).
_NAME_AMT_RE = re.compile(
    r'^(.+?)\s+([\-\(]?[\d,]+\.\d{2}[\)]?)\s*$'
)


def _clean(name: str) -> str:
    return _DOTS_RE.sub('', name).strip()


def _col_has(cell: str, keywords: List[str]) -> bool:
    cell = cell.strip()
    for kw in keywords:
        if kw == cell or kw in cell:
            return True
    return False


def _is_skip(name: str) -> bool:
    nl = name.lower()
    return any(kw in nl for kw in _SKIP_ROW_KW)


# ── TrialBalanceEntry ─────────────────────────────────────────────────────────

class TrialBalanceEntry:
    __slots__ = ('account_name', 'group', 'opening_dr', 'opening_cr',
                 'debit', 'credit', 'closing_dr', 'closing_cr',
                 'is_subtotal', 'is_total')

    def __init__(self, account_name: str, group: str = ''):
        self.account_name = account_name
        self.group = group
        self.opening_dr: float = 0.0
        self.opening_cr: float = 0.0
        self.debit: float = 0.0
        self.credit: float = 0.0
        self.closing_dr: float = 0.0
        self.closing_cr: float = 0.0
        self.is_subtotal: bool = False
        self.is_total: bool = False

    def __repr__(self):
        return f"TBEntry({self.account_name!r} Dr={self.debit} Cr={self.credit})"


# ── Column detection (Format A) ───────────────────────────────────────────────

def _detect_columns(df: pd.DataFrame) -> Tuple[int, Dict[str, int]]:
    for row_idx in range(min(20, len(df))):
        row = [str(c).lower().strip() for c in df.iloc[row_idx]]
        col_map: Dict[str, int] = {}
        for ci, cell in enumerate(row):
            if not cell or cell in ('nan',):
                continue
            if 'account' not in col_map and _col_has(cell, _ACCOUNT_KEYWORDS):
                col_map['account'] = ci
            if 'debit' not in col_map and _col_has(cell, _DEBIT_KEYWORDS):
                col_map['debit'] = ci
            if 'credit' not in col_map and _col_has(cell, _CREDIT_KEYWORDS):
                col_map['credit'] = ci
            if 'open' in cell and 'dr' in cell:
                col_map['opening_dr'] = ci
            if 'open' in cell and 'cr' in cell:
                col_map['opening_cr'] = ci
            if 'clos' in cell and 'dr' in cell:
                col_map['closing_dr'] = ci
            if 'clos' in cell and 'cr' in cell:
                col_map['closing_cr'] = ci

        if 'debit' in col_map and 'credit' in col_map:
            col_map.setdefault('account', 0)
            logger.debug("TB header at row %d: %s", row_idx, col_map)
            return row_idx, col_map
    return -1, {}


# ── Format B: Liabilities/Assets header (UGARARAM style) ─────────────────────

def _is_balance_sheet_format(df: pd.DataFrame) -> bool:
    for row_idx in range(min(12, len(df))):
        row = [str(c).lower().strip() for c in df.iloc[row_idx]]
        non_empty = [c for c in row if c and c != 'nan']
        if not non_empty:
            continue
        has_liab = any(_col_has(c, _LIAB_KW) for c in non_empty)
        has_asset = any(_col_has(c, _ASSET_KW) for c in non_empty)
        if has_liab and has_asset:
            return True
    return False


def _find_bs_header_row(df: pd.DataFrame) -> Tuple[int, int, int, int, int, int, int]:
    """
    Find header row and numeric column positions for a two-sided BS/P&L.
    Scans data rows to locate actual numeric columns on each side.
    Handles non-adjacent layouts (e.g. UGARARAM: name@col1, amounts@col2/3 left;
    name@col5, amounts@col9/11 right).
    """
    _NUM_RE = re.compile(r'^-?[\d,]+\.?\d*$')

    def _is_numeric(v):
        v = str(v).strip()
        return bool(v and v not in ('nan', '0', '') and _NUM_RE.match(v.replace(',', '')))

    for row_idx in range(min(15, len(df))):
        row_raw = [str(c).strip() for c in df.iloc[row_idx]]
        row_low = [c.lower() for c in row_raw]
        non_empty_low = [c for c in row_low if c and c != 'nan']
        if not non_empty_low:
            continue
        has_liab = any(_col_has(c, _LIAB_KW) for c in non_empty_low)
        has_asset = any(_col_has(c, _ASSET_KW) for c in non_empty_low)
        if not (has_liab and has_asset):
            continue

        left_name_col = right_name_col = -1
        for ci, c in enumerate(row_low):
            if _col_has(c, _LIAB_KW) and left_name_col < 0:
                left_name_col = ci
            elif _col_has(c, _ASSET_KW) and right_name_col < 0:
                right_name_col = ci

        if left_name_col < 0:
            left_name_col = 0
        if right_name_col < 0:
            right_name_col = 3

        mid_col = (left_name_col + right_name_col) // 2 + 1
        ncols = df.shape[1]

        # Scan data rows to find actual numeric column positions
        left_numeric_cols = []
        right_numeric_cols = []
        for scan_row in range(row_idx + 1, min(row_idx + 20, len(df))):
            row_s = [str(c).strip() for c in df.iloc[scan_row]]
            for ci in range(left_name_col + 1, min(mid_col, ncols)):
                if _is_numeric(row_s[ci] if ci < len(row_s) else '') and ci not in left_numeric_cols:
                    left_numeric_cols.append(ci)
            for ci in range(right_name_col + 1, ncols):
                if _is_numeric(row_s[ci] if ci < len(row_s) else '') and ci not in right_numeric_cols:
                    right_numeric_cols.append(ci)

        left_numeric_cols.sort()
        right_numeric_cols.sort()

        left_sub_col    = left_numeric_cols[0]  if len(left_numeric_cols) >= 1 else left_name_col + 1
        left_total_col  = left_numeric_cols[1]  if len(left_numeric_cols) >= 2 else left_name_col + 2
        right_sub_col   = right_numeric_cols[0] if len(right_numeric_cols) >= 1 else right_name_col + 1
        right_total_col = right_numeric_cols[1] if len(right_numeric_cols) >= 2 else right_name_col + 2

        logger.debug(
            "BS header row=%d Lcols=(%d,%d,%d) Rcols=(%d,%d,%d)",
            row_idx, left_name_col, left_sub_col, left_total_col,
            right_name_col, right_sub_col, right_total_col
        )
        return row_idx, left_name_col, left_sub_col, left_total_col, right_name_col, right_sub_col, right_total_col

    return -1, 0, 1, 2, 3, 4, 5


def _parse_balance_sheet_format(df: pd.DataFrame) -> List[TrialBalanceEntry]:
    result = _find_bs_header_row(df)
    if result[0] < 0:
        return []
    hdr_row, ln_col, ls_col, lt_col, rn_col, rs_col, rt_col = result

    # Detect P&L vs Balance Sheet for correct Dr/Cr assignment
    is_pl = _is_pl_document(df, hdr_row)

    entries: List[TrialBalanceEntry] = []
    left_group = ''
    right_group = ''

    def safe_col(row_vals, col):
        return row_vals[col] if 0 <= col < len(row_vals) else ''

    for row_idx in range(hdr_row + 1, len(df)):
        row_raw = [str(c).strip() for c in df.iloc[row_idx]]

        # ── Left side ─────────────────────────────────────────────────────
        # BS: Liabilities → Credit   |   P&L: Expenses → Debit
        left_name  = _clean(safe_col(row_raw, ln_col))
        left_sub   = parse_amount(safe_col(row_raw, ls_col))
        left_total = parse_amount(safe_col(row_raw, lt_col))

        if left_name and left_name.lower() not in ('nan', 'liabilit', 'liabilities',
                                                     'liablities', 'amount', ''):
            if _is_skip(left_name):
                pass
            elif left_sub == 0.0 and left_total != 0.0:
                left_group = left_name
            elif left_sub != 0.0:
                amt = left_sub
                e = TrialBalanceEntry(account_name=left_name, group=left_group)
                if is_pl:
                    # P&L left = Debit (Expenses / Purchases)
                    e.debit  = abs(amt) if amt < 0 else amt
                    e.credit = 0.0
                    if amt < 0:
                        e.debit, e.credit = 0.0, abs(amt)
                else:
                    # BS left = Credit (Liabilities)
                    if amt < 0:
                        e.debit = abs(amt)
                    else:
                        e.credit = amt
                entries.append(e)

        # ── Right side ────────────────────────────────────────────────────
        # BS: Assets → Debit   |   P&L: Income → Credit
        right_name  = _clean(safe_col(row_raw, rn_col))
        right_sub   = parse_amount(safe_col(row_raw, rs_col))
        right_total = parse_amount(safe_col(row_raw, rt_col))

        if right_name and right_name.lower() not in ('nan', 'asset', 'assets', 'amount', ''):
            if _is_skip(right_name):
                pass
            elif right_sub == 0.0 and right_total != 0.0:
                right_group = right_name
            elif right_sub != 0.0:
                amt = right_sub
                e = TrialBalanceEntry(account_name=right_name, group=right_group)
                if is_pl:
                    # P&L right = Credit (Income / Sales)
                    if amt < 0:
                        e.debit  = abs(amt)
                    else:
                        e.credit = amt
                else:
                    # BS right = Debit (Assets)
                    if amt < 0:
                        e.credit = abs(amt)
                    else:
                        e.debit = amt
                entries.append(e)

    return entries


# ── Format C: PARTICULARS/AMOUNT two-column XLS (UMA TEXCOM Type 1) ──────────
#
# Structure (10 cols):
#   col 0: left group header (when non-blank, col 4 = group total)
#   col 1: left leaf account name (indented sub-entry)
#   col 2: (unused / merged cell artifact)
#   col 3: left leaf amount
#   col 4: left group total (non-blank only on group-header rows)
#   col 5: right group header
#   col 6: right leaf account name
#   col 7: (unused)
#   col 8: right leaf amount
#   col 9: right group total
#
# Left side → Credit (Liabilities/Expenses)
# Right side → Debit (Assets/Income)

def _is_particulars_two_column(df: pd.DataFrame) -> bool:
    """Detect PARTICULARS | AMOUNT Rs. | PARTICULARS | AMOUNT RS. header."""
    for row_idx in range(min(15, len(df))):
        row = [str(c).lower().strip() for c in df.iloc[row_idx]]
        non_nan = [c for c in row if c and c != 'nan']
        if not non_nan:
            continue
        particulars_count = sum(1 for c in non_nan if 'particulars' in c)
        amount_count = sum(1 for c in non_nan if 'amount' in c)
        if particulars_count >= 2 and amount_count >= 2:
            return True
        # Also catch single PARTICULARS + two AMOUNT headers
        if particulars_count >= 1 and amount_count >= 2:
            return True
    return False


def _find_particulars_header_row(df: pd.DataFrame) -> int:
    for row_idx in range(min(15, len(df))):
        row = [str(c).lower().strip() for c in df.iloc[row_idx]]
        non_nan = [c for c in row if c and c != 'nan']
        if not non_nan:
            continue
        particulars_count = sum(1 for c in non_nan if 'particulars' in c)
        amount_count = sum(1 for c in non_nan if 'amount' in c)
        if (particulars_count >= 2 and amount_count >= 2) or \
           (particulars_count >= 1 and amount_count >= 2):
            return row_idx
    return -1


def _is_pl_document(df: pd.DataFrame, hdr_row: int) -> bool:
    """Return True if the document is a P&L / Trading account (not a Balance Sheet)."""
    _PL_KW = ['profit', 'loss', 'trading', 'p & l', 'p&l', 'income', 'expenditure']
    for ri in range(min(hdr_row, 15)):
        row_text = ' '.join(str(c).lower().strip() for c in df.iloc[ri]
                            if str(c).strip() not in ('nan', ''))
        if any(kw in row_text for kw in _PL_KW):
            return True
    return False


def _parse_particulars_two_column(df: pd.DataFrame) -> List[TrialBalanceEntry]:
    """
    Parse UMA TEXCOM two-column P&L / Balance Sheet XLS.

    10-column layout — two side-by-side tables:
      Left  (col0=group, col1=leaf_name, col3=amount_A, col4=amount_B)
      Right (col5=group, col6=leaf_name, col8=amount_A, col9=amount_B)

    Amount column convention varies by document type:
      Balance Sheet: group-total in col4, leaf-amount in col3
      P&L Account:   group-total in col3, leaf-amount in col4
    We auto-detect by checking which column has amounts on indented (col1) rows.

    SIDE ASSIGNMENT (critical):
      Balance Sheet  → Left=Credit (Liabilities),  Right=Debit (Assets)
      P&L / Trading  → Left=Debit  (Expenses),     Right=Credit (Income)
    """
    hdr_row = _find_particulars_header_row(df)
    if hdr_row < 0:
        return []

    # Detect if this is a P&L document (left=Debit) or BS (left=Credit)
    is_pl = _is_pl_document(df, hdr_row)
    logger.debug("Format C: is_pl=%s (hdr_row=%d)", is_pl, hdr_row)

    def sv(row_vals, col):
        if 0 <= col < len(row_vals):
            v = str(row_vals[col]).strip()
            return '' if v == 'nan' else v
        return ''

    # Auto-detect: scan first 15 data rows to find which col (3 or 4) has amounts
    # on rows where col1 is non-empty (leaf rows). Whichever has more non-zero
    # amounts is the leaf-amount column; the other is the group-total column.
    col3_hits = col4_hits = 0
    for ri in range(hdr_row + 1, min(hdr_row + 20, len(df))):
        row = list(df.iloc[ri])
        if sv(row, 1) and sv(row, 1) not in ('nan', ''):
            if parse_amount(sv(row, 3)) != 0.0: col3_hits += 1
            if parse_amount(sv(row, 4)) != 0.0: col4_hits += 1
    # Leaf col is where amounts appear on indented (col1) rows
    if col3_hits >= col4_hits:
        left_leaf_col, left_grp_col   = 3, 4
    else:
        left_leaf_col, left_grp_col   = 4, 3

    # Same detection for right side (cols 8, 9)
    col8_hits = col9_hits = 0
    for ri in range(hdr_row + 1, min(hdr_row + 20, len(df))):
        row = list(df.iloc[ri])
        if sv(row, 6) and sv(row, 6) not in ('nan', ''):
            if parse_amount(sv(row, 8)) != 0.0: col8_hits += 1
            if parse_amount(sv(row, 9)) != 0.0: col9_hits += 1
    if col8_hits >= col9_hits:
        right_leaf_col, right_grp_col = 8, 9
    else:
        right_leaf_col, right_grp_col = 9, 8

    logger.debug("Format C cols: left_leaf=%d left_grp=%d right_leaf=%d right_grp=%d",
                 left_leaf_col, left_grp_col, right_leaf_col, right_grp_col)

    entries: List[TrialBalanceEntry] = []
    left_group = ''
    right_group = ''

    for row_idx in range(hdr_row + 1, len(df)):
        row_raw = list(df.iloc[row_idx])

        # ── Left side ──────────────────────────────────────────────────────
        l0        = _clean(sv(row_raw, 0))
        l1        = _clean(sv(row_raw, 1))
        l_leaf    = parse_amount(sv(row_raw, left_leaf_col))
        l_grp_tot = parse_amount(sv(row_raw, left_grp_col))

        if l0 and not _is_skip(l0):
            # Group header: col0 has name, grp_col has total, leaf_col is 0
            if l_grp_tot != 0.0 and l_leaf == 0.0:
                left_group = l0
            elif l_leaf != 0.0:
                # Some formats put group total in leaf col — treat as group header
                left_group = l0
                # Don't create an entry for the group header itself (it's a summary)

        if l1 and not _is_skip(l1) and l_leaf != 0.0:
            e = TrialBalanceEntry(account_name=l1, group=left_group)
            if is_pl:
                # P&L left side = Debit (Expenses / Purchases)
                if l_leaf < 0:
                    e.credit = abs(l_leaf)
                else:
                    e.debit = l_leaf
            else:
                # BS left side = Credit (Liabilities)
                if l_leaf < 0:
                    e.debit  = abs(l_leaf)
                else:
                    e.credit = l_leaf
            entries.append(e)

        # ── Right side ─────────────────────────────────────────────────────
        r5        = _clean(sv(row_raw, 5))
        r6        = _clean(sv(row_raw, 6))
        r_leaf    = parse_amount(sv(row_raw, right_leaf_col))
        r_grp_tot = parse_amount(sv(row_raw, right_grp_col))

        if r5 and not _is_skip(r5):
            if r_grp_tot != 0.0 and r_leaf == 0.0:
                right_group = r5
            elif r_leaf != 0.0:
                right_group = r5

        if r6 and not _is_skip(r6) and r_leaf != 0.0:
            e = TrialBalanceEntry(account_name=r6, group=right_group)
            if is_pl:
                # P&L right side = Credit (Income / Sales)
                if r_leaf < 0:
                    e.debit  = abs(r_leaf)
                else:
                    e.credit = r_leaf
            else:
                # BS right side = Debit (Assets)
                if r_leaf < 0:
                    e.credit = abs(r_leaf)
                else:
                    e.debit  = r_leaf
            entries.append(e)

    return entries


# ── Format D: PDF two-column cell parsing ─────────────────────────────────────
#
# pdfplumber sees the PDF as a 2-column table:
#   col 0 = left side text block: "ACCOUNT NAME  AMOUNT\nSUB NAME  AMOUNT\n..."
#   col 1 = right side text block: "ACCOUNT NAME  AMOUNT\n..."
#
# Each column cell is a multi-line string with "NAME  AMOUNT" on each line.
# Left side → Credit, Right side → Debit

def _parse_name_amount_lines(text: str) -> List[Tuple[str, float]]:
    """Extract (name, amount) pairs from a multi-line PDF cell text block."""
    _skip = {'nan', 'particulars', 'amount rs.', 'amount rs', 'amount', '', 'total', 'grand total'}
    results = []
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        m = _NAME_AMT_RE.match(line)
        if m:
            name = _clean(m.group(1).strip())
            amt_str = m.group(2).strip()
            if amt_str.startswith('(') and amt_str.endswith(')'):
                amt_str = '-' + amt_str[1:-1]
            amt = parse_amount(amt_str)
            if name and name.lower() not in _skip:
                results.append((name, amt))
    return results


def _detect_pdf_groups(pairs: List[Tuple[str, float]]) -> List[Tuple[str, float, str]]:
    """
    Given a flat list of (name, amount) pairs from one PDF column,
    detect group HEADER lines (whose amount = sum of following sub-entries)
    and return (name, amount, group_name) for every LEAF entry.

    Algorithm:
      For each entry i, check if its absolute amount equals the sum of absolute
      amounts of some consecutive entries starting at i+1 (within ±0.10 tolerance).
      We look ahead up to 500 entries to cover large groups (e.g. SUNDRY CREDITORS
      with 280+ sub-entries).  Mixed-sign groups (some sub-entries negative) are
      handled by summing absolute values.
    """
    n = len(pairs)
    is_header = [False] * n

    for i in range(n):
        name, amt = pairs[i]
        abs_amt = abs(amt)
        if abs_amt < 0.01:
            continue
        # Try signed sum first — handles groups where some sub-items are negative
        # (e.g. SALE A/C has CLOTH SALE + GOOD RETURN (negative) + YARN SALE etc.)
        # Also try absolute-value sum as fallback for uniformly-signed groups.
        running_sign = 0.0
        running_abs  = 0.0
        found = False
        for j in range(i + 1, min(i + 500, n)):
            running_sign += pairs[j][1]
            running_abs  += abs(pairs[j][1])
            match_sign = abs(running_sign - amt)  < 0.50
            match_abs  = abs(running_abs  - abs_amt) < 0.50
            if match_sign or match_abs:
                is_header[i] = True
                found = True
                break
            # Only abort early if the unsigned running sum is far past abs_amt
            # AND signed sum is also far past — meaning no sub-entry can bring it back
            if running_abs > abs_amt * 4 + 1000:
                break
        _ = found  # suppress unused warning

    results = []
    current_group = ''
    for i, (name, amt) in enumerate(pairs):
        if _is_skip(name):
            continue
        if is_header[i]:
            current_group = name
        else:
            results.append((name, amt, current_group))

    return results


def _parse_pdf_two_column_table(table) -> List[TrialBalanceEntry]:
    """
    Parse a pdfplumber 2-column table (left = credit/liability side, right = debit/asset side).
    Collects ALL text from each side across ALL pages, then runs group detection once
    so that group totals spanning many sub-entries are correctly identified.
    """
    # Accumulate all (name, amount) pairs for each side across the whole table
    all_left_pairs:  List[Tuple[str, float]] = []
    all_right_pairs: List[Tuple[str, float]] = []

    for row in table:
        if not row or len(row) < 2:
            continue
        left_cell  = str(row[0] or '')
        right_cell = str(row[1] or '')

        # Skip header rows
        left_low = left_cell.lower().strip()
        if 'particulars' in left_low and len(left_low) < 60:
            continue

        all_left_pairs.extend(_parse_name_amount_lines(left_cell))
        all_right_pairs.extend(_parse_name_amount_lines(right_cell))

    entries: List[TrialBalanceEntry] = []

    # Left side → Credit (liability/expense)
    for name, amt, grp in _detect_pdf_groups(all_left_pairs):
        e = TrialBalanceEntry(account_name=name, group=grp)
        if amt < 0:
            e.debit  = abs(amt)
        else:
            e.credit = amt
        entries.append(e)

    # Right side → Debit (asset/income)
    for name, amt, grp in _detect_pdf_groups(all_right_pairs):
        e = TrialBalanceEntry(account_name=name, group=grp)
        if amt < 0:
            e.credit = abs(amt)
        else:
            e.debit  = amt
        entries.append(e)

    return entries


# ── Main parser ───────────────────────────────────────────────────────────────

class TrialBalanceParser:

    def parse_file(self, file_path: str) -> List[TrialBalanceEntry]:
        ext = file_path.rsplit('.', 1)[-1].lower()

        if ext == 'pdf':
            return self._parse_pdf(file_path)

        try:
            if ext in ('xlsx', 'xls'):
                engine = 'xlrd' if ext == 'xls' else 'openpyxl'
                df = pd.read_excel(file_path, header=None, dtype=str, engine=engine)
            elif ext == 'csv':
                df = pd.read_csv(file_path, header=None, dtype=str)
            else:
                raise UnsupportedFileTypeError(f"Unsupported file type: {ext}")
        except UnsupportedFileTypeError:
            raise
        except Exception as e:
            raise FileParseError(f"Failed to read file: {e}")

        df = df.fillna('').map(lambda x: str(x).strip())
        return self._parse_dataframe(df, file_path)

    def _parse_dataframe(self, df: pd.DataFrame, source: str = '') -> List[TrialBalanceEntry]:
        # Strategy A: Standard Dr/Cr header detection
        header_row, col_map = _detect_columns(df)
        if header_row >= 0:
            entries = self._extract_standard(df, header_row, col_map)
            if entries:
                logger.info("Parsed %d TB entries (standard format) from %s", len(entries), source)
                return entries

        # Strategy B: Balance Sheet / P&L two-sided layout with Liabilities/Assets header
        if _is_balance_sheet_format(df):
            logger.info("Detected Balance Sheet / P&L format (Liabilities/Assets) in %s", source)
            entries = _parse_balance_sheet_format(df)
            if entries:
                logger.info("Parsed %d TB entries (BS/PL Liab/Asset format) from %s", len(entries), source)
                return entries

        # Strategy C: Two-sided P&L/BS with PARTICULARS + AMOUNT headers (UMA TEXCOM XLS style)
        if _is_particulars_two_column(df):
            logger.info("Detected two-column PARTICULARS/AMOUNT format in %s", source)
            entries = _parse_particulars_two_column(df)
            if entries:
                logger.info("Parsed %d TB entries (PARTICULARS two-col format) from %s", len(entries), source)
                return entries

        raise FileParseError(
            "Could not detect column headers (Debit/Credit). "
            "Ensure the trial balance file has standard column headers.\n"
            "Supported formats:\n"
            "  • Standard TB: 'Amount (Dr)'/'Amount (Cr)' or 'Debit'/'Credit' or 'Dr'/'Cr'\n"
            "  • Balance Sheet/P&L (Liabilities/Assets headers)\n"
            "  • Balance Sheet/P&L (PARTICULARS/AMOUNT two-column layout)"
        )

    def _extract_standard(
        self, df: pd.DataFrame, header_row: int, col_map: Dict[str, int]
    ) -> List[TrialBalanceEntry]:
        entries = []
        current_group = ''

        for row_idx in range(header_row + 1, len(df)):
            row = df.iloc[row_idx]
            account_val = _clean(str(row.iloc[col_map['account']]))
            debit_val   = str(row.iloc[col_map['debit']]).strip()
            credit_val  = str(row.iloc[col_map['credit']]).strip()

            if not account_val and not debit_val and not credit_val:
                continue

            debit  = parse_amount(debit_val)
            credit = parse_amount(credit_val)

            if _is_skip(account_val):
                continue

            if account_val and debit == 0.0 and credit == 0.0:
                current_group = account_val
                continue

            if not account_val:
                continue

            e = TrialBalanceEntry(account_name=account_val, group=current_group)
            e.debit  = debit
            e.credit = credit

            if 'opening_dr' in col_map:
                e.opening_dr = parse_amount(str(row.iloc[col_map['opening_dr']]))
            if 'opening_cr' in col_map:
                e.opening_cr = parse_amount(str(row.iloc[col_map['opening_cr']]))
            if 'closing_dr' in col_map:
                e.closing_dr = parse_amount(str(row.iloc[col_map['closing_dr']]))
            if 'closing_cr' in col_map:
                e.closing_cr = parse_amount(str(row.iloc[col_map['closing_cr']]))

            entries.append(e)
        return entries

    # ── PDF ───────────────────────────────────────────────────────────────────

    def _parse_pdf(self, file_path: str) -> List[TrialBalanceEntry]:
        import pdfplumber

        all_entries: List[TrialBalanceEntry] = []

        # ── Pass 1: pdfplumber table extraction ──────────────────────────────
        # For two-column PARTICULARS/AMOUNT PDFs we must collect ALL pages' pairs
        # before running group detection, since groups can span multiple pages.
        try:
            two_col_left_pairs:  List[Tuple[str, float]] = []
            two_col_right_pairs: List[Tuple[str, float]] = []
            is_two_col_pdf = False
            standard_table_entries: List[TrialBalanceEntry] = []

            with pdfplumber.open(file_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    tables = page.extract_tables() or []
                    for table in tables:
                        if not table or len(table) < 2:
                            continue

                        # Detect two-column PARTICULARS/AMOUNT layout
                        if len(table[0]) == 2:
                            header_text = ' '.join(str(c or '').lower() for c in table[0])
                            if 'particulars' in header_text or 'amount' in header_text:
                                is_two_col_pdf = True
                                # Accumulate pairs from ALL pages
                                for row in table:
                                    if not row or len(row) < 2:
                                        continue
                                    left_cell  = str(row[0] or '')
                                    right_cell = str(row[1] or '')
                                    left_low   = left_cell.lower().strip()
                                    if 'particulars' in left_low and len(left_low) < 60:
                                        continue
                                    two_col_left_pairs.extend(_parse_name_amount_lines(left_cell))
                                    two_col_right_pairs.extend(_parse_name_amount_lines(right_cell))
                                continue

                        # Standard Dr/Cr table detection
                        for hdr_idx, row in enumerate(table[:5]):
                            if not row:
                                continue
                            hdr_low = [str(c or '').lower().strip() for c in row]
                            if (any(_col_has(c, _DEBIT_KEYWORDS) for c in hdr_low) and
                                    any(_col_has(c, _CREDIT_KEYWORDS) for c in hdr_low)):
                                entries = self._extract_from_pdf_table(table, hdr_idx, hdr_low)
                                if entries:
                                    standard_table_entries.extend(entries)
                                    break

            if is_two_col_pdf and (two_col_left_pairs or two_col_right_pairs):
                # Run group detection once across all pages combined
                for name, amt, grp in _detect_pdf_groups(two_col_left_pairs):
                    e = TrialBalanceEntry(account_name=name, group=grp)
                    if amt < 0:
                        e.debit  = abs(amt)
                    else:
                        e.credit = amt
                    all_entries.append(e)

                for name, amt, grp in _detect_pdf_groups(two_col_right_pairs):
                    e = TrialBalanceEntry(account_name=name, group=grp)
                    if amt < 0:
                        e.credit = abs(amt)
                    else:
                        e.debit  = amt
                    all_entries.append(e)

                logger.info("PDF (two-col): %d entries, %d left pairs, %d right pairs",
                            len(all_entries), len(two_col_left_pairs), len(two_col_right_pairs))

            if standard_table_entries:
                all_entries.extend(standard_table_entries)

            if all_entries:
                logger.info("PDF parsed %d total entries via table extraction", len(all_entries))
                return all_entries

        except Exception as e:
            logger.warning("pdfplumber table extraction failed: %s", e)

        # ── Pass 2: Text-line fallback ───────────────────────────────────────
        raw_lines = []
        for lib in ['pypdf', 'pdfplumber']:
            try:
                if lib == 'pypdf':
                    from pypdf import PdfReader
                    reader = PdfReader(file_path)
                    for page in reader.pages:
                        for line in (page.extract_text() or '').split('\n'):
                            line = line.strip()
                            if line:
                                raw_lines.append(re.split(r'\s{2,}', line))
                else:
                    with pdfplumber.open(file_path) as pdf:
                        for page in pdf.pages:
                            for line in (page.extract_text() or '').split('\n'):
                                line = line.strip()
                                if line:
                                    raw_lines.append(re.split(r'\s{2,}', line))
                if raw_lines:
                    break
            except Exception:
                continue

        if not raw_lines:
            raise EmptyFileError("No text found in PDF.")

        max_cols = max(len(r) for r in raw_lines)
        padded = [r + [''] * (max_cols - len(r)) for r in raw_lines]
        df = pd.DataFrame(padded, dtype=str).fillna('').map(lambda x: str(x).strip())

        # Try all dataframe strategies
        try:
            return self._parse_dataframe(df, file_path)
        except FileParseError:
            pass

        # ── Pass 3: Text-line two-column heuristic ───────────────────────────
        # For PDFs where text extraction gives "NAME  AMOUNT  NAME  AMOUNT" lines
        entries = self._parse_pdf_text_two_column(raw_lines)
        if entries:
            logger.info("PDF parsed %d entries via text two-column heuristic", len(entries))
            return entries

        raise FileParseError(
            "Could not detect Debit/Credit columns in the PDF. "
            "Ensure the trial balance PDF has standard column headers such as "
            "'Debit'/'Credit', 'Dr'/'Cr', 'Amount (Dr)'/'Amount (Cr)', "
            "or 'Liabilities'/'Assets'."
        )

    def _parse_pdf_text_two_column(
        self, raw_lines: List[List[str]]
    ) -> List[TrialBalanceEntry]:
        """
        Heuristic for PDFs whose text extraction produces lines like:
          ['CAPITAL A/C  8841271.50', 'LOANS & ADVANCES  9349103.00']
        where col 0 = left side (credit), col 1 = right side (debit).
        """
        # Collect all pairs from each side first, then run group detection
        all_left_pairs:  List[Tuple[str, float]] = []
        all_right_pairs: List[Tuple[str, float]] = []

        for parts in raw_lines:
            if len(parts) < 2:
                continue
            left_text  = parts[0].strip()
            right_text = parts[1].strip()
            if any(kw in left_text.lower() for kw in ('particulars', 'amount rs', 'page no')):
                continue
            all_left_pairs.extend(_parse_name_amount_lines(left_text))
            all_right_pairs.extend(_parse_name_amount_lines(right_text))

        entries: List[TrialBalanceEntry] = []

        for name, amt, grp in _detect_pdf_groups(all_left_pairs):
            e = TrialBalanceEntry(account_name=name, group=grp)
            e.credit = amt if amt > 0 else 0.0
            e.debit  = abs(amt) if amt < 0 else 0.0
            entries.append(e)

        for name, amt, grp in _detect_pdf_groups(all_right_pairs):
            e = TrialBalanceEntry(account_name=name, group=grp)
            e.debit  = amt if amt > 0 else 0.0
            e.credit = abs(amt) if amt < 0 else 0.0
            entries.append(e)

        return entries

    def _extract_from_pdf_table(
        self, table, hdr_idx: int, hdr_low: List[str]
    ) -> List[TrialBalanceEntry]:
        debit_col = credit_col = account_col = -1
        for i, cell in enumerate(hdr_low):
            if debit_col < 0 and _col_has(cell, _DEBIT_KEYWORDS):
                debit_col = i
            elif credit_col < 0 and _col_has(cell, _CREDIT_KEYWORDS):
                credit_col = i
            elif account_col < 0 and _col_has(cell, _ACCOUNT_KEYWORDS):
                account_col = i
        if account_col < 0:
            account_col = 0
        if debit_col < 0 or credit_col < 0:
            return []

        entries = []
        current_group = ''
        for row in table[hdr_idx + 1:]:
            if not row:
                continue
            row = [str(c or '').strip() for c in row]
            account = _clean(row[account_col] if account_col < len(row) else '')
            debit   = parse_amount(row[debit_col]  if debit_col  < len(row) else '')
            credit  = parse_amount(row[credit_col] if credit_col < len(row) else '')
            if not account or account.lower() in ('nan', ''):
                continue
            if _is_skip(account):
                continue
            if debit == 0.0 and credit == 0.0:
                current_group = account
                continue
            e = TrialBalanceEntry(account_name=account, group=current_group)
            e.debit  = debit
            e.credit = credit
            entries.append(e)
        return entries
