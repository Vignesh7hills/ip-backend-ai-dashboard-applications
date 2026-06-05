"""
Trial Balance Parser — ROBUST v3.

Supports all known input formats:

  Format A — Standard Dr/Cr TB:
    Sr No | Ledger Name | Group | Amount (Dr) | Amount (Cr)

  Format B — Two-sided BS/P&L with Liabilities/Assets header (UGARARAM):
    Liabilities | Amount | Assets | Amount

  Format C — PARTICULARS/AMOUNT two-column (UMA TEXCOM XLS):
    PARTICULARS | AMOUNT Rs. | PARTICULARS | AMOUNT Rs.

  Format D — PDF two-column table (UMA TEXCOM bs.pdf / pl.pdf):
    pdfplumber extracts as 2-column: left_text | right_text

  Format E — Single-column Dr/Cr with inline direction (new):
    Account Name | Amount (may have Cr/Dr suffix) | Group

Improvements over v2:
  - Smarter column detection: validates column content, not just headers
  - Handles amount columns that have "Dr"/"Cr" suffix in header or in data
  - Handles multi-row headers (merged cells in Excel)
  - Better group detection for PDF two-column tables
  - Tolerant of rows with partial data (skips gracefully)
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
    'debit balance', 'opening dr', 'closing dr', 'dr bal',
    'debit', 'amount dr', 'dr.', 'dr',
]
_CREDIT_KEYWORDS = [
    'amount (cr)', 'amount(cr)', 'cr amount', 'credit amount',
    'credit balance', 'opening cr', 'closing cr', 'cr bal',
    'credit', 'amount cr', 'cr.', 'cr',
]
_ACCOUNT_KEYWORDS = [
    'ledger name', 'account name', 'account head', 'particulars',
    'description', 'narration', 'ledger', 'account', 'head', 'name',
]
_LIAB_KW  = ['liabilit', 'liabilities', 'liablities']
_ASSET_KW = ['asset', 'assets']

_SKIP_ROW_KW = {
    'total', 'grand total', 'sub total', 'subtotal', 'net total',
    'gross profit', 'net profit', 'net loss',
}

_DOTS_RE   = re.compile(r'(\s*\.\s*){2,}')
_NUM_RE    = re.compile(r'^-?[\d,]+\.?\d*$')
_NAME_AMT_RE = re.compile(r'^(.+?)\s+([\-\(]?[\d,]+\.\d{2}[\)]?)\s*$')


def _clean(name: str) -> str:
    return _DOTS_RE.sub('', name).strip()


def _col_has(cell: str, keywords: List[str]) -> bool:
    cell_l = cell.strip().lower()
    for kw in keywords:
        if kw == cell_l or kw in cell_l:
            return True
    return False


def _is_skip(name: str) -> bool:
    nl = name.lower().strip()
    return any(kw in nl for kw in _SKIP_ROW_KW)


def _is_numeric(v) -> bool:
    v = str(v).strip()
    return bool(v and v not in ('nan', '0', '') and _NUM_RE.match(v.replace(',', '')))


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


# ── Merge multi-row headers ────────────────────────────────────────────────────

def _merge_header_rows(df: pd.DataFrame, max_rows: int = 4) -> pd.DataFrame:
    """
    Merge consecutive text-only rows into one combined header row.
    Handles Excel files where column names span 2-3 rows.
    """
    def _looks_header(ri):
        row = [str(c).strip() for c in df.iloc[ri]]
        non_empty = [c for c in row if c and c.lower() not in ('nan', '')]
        if not non_empty:
            return True  # blank row — skip it
        nums = sum(1 for c in non_empty if _is_numeric(c))
        return nums < len(non_empty) * 0.5

    merge_n = 0
    for ri in range(min(max_rows, len(df))):
        if _looks_header(ri):
            merge_n += 1
        else:
            break

    if merge_n <= 1:
        return df

    merged = []
    for ci in range(df.shape[1]):
        parts = []
        for ri in range(merge_n):
            v = str(df.iloc[ri, ci]).strip()
            if v and v.lower() not in ('nan', ''):
                parts.append(v)
        merged.append(' '.join(parts))

    merged_row = pd.DataFrame([merged], columns=df.columns)
    rest = df.iloc[merge_n:].reset_index(drop=True)
    return pd.concat([merged_row, rest], ignore_index=True)


# ── Format A: Standard Dr/Cr column detection ─────────────────────────────────

def _detect_columns(df: pd.DataFrame) -> Tuple[int, Dict[str, int]]:
    df2 = _merge_header_rows(df)
    for row_idx in range(min(20, len(df2))):
        row = [str(c).lower().strip() for c in df2.iloc[row_idx]]
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
            # Validate: check actual data rows have numbers in these columns
            dr_col = col_map['debit']
            cr_col = col_map['credit']
            has_data = False
            for check_ri in range(row_idx + 1, min(row_idx + 10, len(df2))):
                dr_v = str(df2.iloc[check_ri, dr_col]).strip()
                cr_v = str(df2.iloc[check_ri, cr_col]).strip()
                if _is_numeric(dr_v) or _is_numeric(cr_v):
                    has_data = True
                    break
            if has_data:
                col_map.setdefault('account', 0)
                # Adjust row_idx back to original df if we merged rows
                offset = len(df) - len(df2)
                logger.debug("TB header at row %d (merged=%d): %s", row_idx, offset, col_map)
                return row_idx, col_map

    return -1, {}


# ── Format B: Liabilities/Assets two-sided ────────────────────────────────────

def _is_balance_sheet_format(df: pd.DataFrame) -> bool:
    for row_idx in range(min(12, len(df))):
        row = [str(c).lower().strip() for c in df.iloc[row_idx]]
        non_empty = [c for c in row if c and c != 'nan']
        if not non_empty:
            continue
        has_liab  = any(_col_has(c, _LIAB_KW)  for c in non_empty)
        has_asset = any(_col_has(c, _ASSET_KW) for c in non_empty)
        if has_liab and has_asset:
            return True
    return False


def _find_bs_header_row(df: pd.DataFrame) -> Tuple[int, int, int, int, int, int, int]:
    for row_idx in range(min(15, len(df))):
        row_raw = [str(c).strip() for c in df.iloc[row_idx]]
        row_low = [c.lower() for c in row_raw]
        non_empty_low = [c for c in row_low if c and c != 'nan']
        if not non_empty_low:
            continue
        has_liab  = any(_col_has(c, _LIAB_KW)  for c in non_empty_low)
        has_asset = any(_col_has(c, _ASSET_KW) for c in non_empty_low)
        if not (has_liab and has_asset):
            continue

        left_name_col = right_name_col = -1
        for ci, c in enumerate(row_low):
            if _col_has(c, _LIAB_KW)  and left_name_col  < 0: left_name_col  = ci
            elif _col_has(c, _ASSET_KW) and right_name_col < 0: right_name_col = ci

        if left_name_col  < 0: left_name_col  = 0
        if right_name_col < 0: right_name_col = 3

        mid_col = (left_name_col + right_name_col) // 2 + 1
        ncols   = df.shape[1]

        left_numeric  = []
        right_numeric = []
        for scan_row in range(row_idx + 1, min(row_idx + 20, len(df))):
            row_s = [str(c).strip() for c in df.iloc[scan_row]]
            for ci in range(left_name_col + 1, min(mid_col, ncols)):
                if _is_numeric(row_s[ci] if ci < len(row_s) else '') and ci not in left_numeric:
                    left_numeric.append(ci)
            for ci in range(right_name_col + 1, ncols):
                if _is_numeric(row_s[ci] if ci < len(row_s) else '') and ci not in right_numeric:
                    right_numeric.append(ci)

        left_numeric.sort()
        right_numeric.sort()

        l_sub  = left_numeric[0]  if len(left_numeric)  >= 1 else left_name_col  + 1
        l_tot  = left_numeric[1]  if len(left_numeric)  >= 2 else left_name_col  + 2
        r_sub  = right_numeric[0] if len(right_numeric) >= 1 else right_name_col + 1
        r_tot  = right_numeric[1] if len(right_numeric) >= 2 else right_name_col + 2

        return row_idx, left_name_col, l_sub, l_tot, right_name_col, r_sub, r_tot

    return -1, 0, 1, 2, 3, 4, 5


def _is_pl_document(df: pd.DataFrame, hdr_row: int) -> bool:
    _PL_KW = ['profit', 'loss', 'trading', 'p & l', 'p&l', 'income', 'expenditure']
    for ri in range(min(hdr_row, 15)):
        row_text = ' '.join(str(c).lower().strip() for c in df.iloc[ri]
                            if str(c).strip() not in ('nan', ''))
        if any(kw in row_text for kw in _PL_KW):
            return True
    return False


def _parse_balance_sheet_format(df: pd.DataFrame) -> List[TrialBalanceEntry]:
    result = _find_bs_header_row(df)
    if result[0] < 0:
        return []
    hdr_row, ln_col, ls_col, lt_col, rn_col, rs_col, rt_col = result
    is_pl = _is_pl_document(df, hdr_row)

    entries: List[TrialBalanceEntry] = []
    left_group = right_group = ''

    def sv(row_vals, col):
        return row_vals[col] if 0 <= col < len(row_vals) else ''

    for row_idx in range(hdr_row + 1, len(df)):
        row_raw = [str(c).strip() for c in df.iloc[row_idx]]

        left_name  = _clean(sv(row_raw, ln_col))
        left_sub   = parse_amount(sv(row_raw, ls_col))
        left_total = parse_amount(sv(row_raw, lt_col))

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
                    e.debit  = abs(amt) if amt >= 0 else 0.0
                    e.credit = abs(amt) if amt < 0  else 0.0
                else:
                    e.credit = amt if amt >= 0 else 0.0
                    e.debit  = abs(amt) if amt < 0  else 0.0
                entries.append(e)

        right_name  = _clean(sv(row_raw, rn_col))
        right_sub   = parse_amount(sv(row_raw, rs_col))
        right_total = parse_amount(sv(row_raw, rt_col))

        if right_name and right_name.lower() not in ('nan', 'asset', 'assets', 'amount', ''):
            if _is_skip(right_name):
                pass
            elif right_sub == 0.0 and right_total != 0.0:
                right_group = right_name
            elif right_sub != 0.0:
                amt = right_sub
                e = TrialBalanceEntry(account_name=right_name, group=right_group)
                if is_pl:
                    e.credit = amt if amt >= 0 else 0.0
                    e.debit  = abs(amt) if amt < 0  else 0.0
                else:
                    e.debit  = amt if amt >= 0 else 0.0
                    e.credit = abs(amt) if amt < 0  else 0.0
                entries.append(e)

    return entries


# ── Format C: PARTICULARS/AMOUNT two-column XLS ───────────────────────────────

def _is_particulars_two_column(df: pd.DataFrame) -> bool:
    for row_idx in range(min(15, len(df))):
        row = [str(c).lower().strip() for c in df.iloc[row_idx]]
        non_nan = [c for c in row if c and c != 'nan']
        if not non_nan:
            continue
        p_count = sum(1 for c in non_nan if 'particulars' in c)
        a_count = sum(1 for c in non_nan if 'amount' in c)
        if p_count >= 2 and a_count >= 2:
            return True
        if p_count >= 1 and a_count >= 2:
            return True
    return False


def _find_particulars_header_row(df: pd.DataFrame) -> int:
    for row_idx in range(min(15, len(df))):
        row = [str(c).lower().strip() for c in df.iloc[row_idx]]
        non_nan = [c for c in row if c and c != 'nan']
        if not non_nan:
            continue
        p_count = sum(1 for c in non_nan if 'particulars' in c)
        a_count = sum(1 for c in non_nan if 'amount' in c)
        if (p_count >= 2 and a_count >= 2) or (p_count >= 1 and a_count >= 2):
            return row_idx
    return -1


def _parse_particulars_two_column(df: pd.DataFrame) -> List[TrialBalanceEntry]:
    hdr_row = _find_particulars_header_row(df)
    if hdr_row < 0:
        return []

    is_pl = _is_pl_document(df, hdr_row)

    def sv(row_vals, col):
        if 0 <= col < len(row_vals):
            v = str(row_vals[col]).strip()
            return '' if v == 'nan' else v
        return ''

    # Auto-detect leaf vs group columns
    col3_hits = col4_hits = 0
    for ri in range(hdr_row + 1, min(hdr_row + 20, len(df))):
        row = list(df.iloc[ri])
        if sv(row, 1) and sv(row, 1) not in ('nan', ''):
            if parse_amount(sv(row, 3)) != 0.0: col3_hits += 1
            if parse_amount(sv(row, 4)) != 0.0: col4_hits += 1
    left_leaf_col  = 3 if col3_hits >= col4_hits else 4
    left_grp_col   = 4 if col3_hits >= col4_hits else 3

    col8_hits = col9_hits = 0
    for ri in range(hdr_row + 1, min(hdr_row + 20, len(df))):
        row = list(df.iloc[ri])
        if sv(row, 6) and sv(row, 6) not in ('nan', ''):
            if parse_amount(sv(row, 8)) != 0.0: col8_hits += 1
            if parse_amount(sv(row, 9)) != 0.0: col9_hits += 1
    right_leaf_col = 8 if col8_hits >= col9_hits else 9
    right_grp_col  = 9 if col8_hits >= col9_hits else 8

    entries: List[TrialBalanceEntry] = []
    left_group = right_group = ''

    for row_idx in range(hdr_row + 1, len(df)):
        row_raw = list(df.iloc[row_idx])

        l0 = _clean(sv(row_raw, 0))
        l1 = _clean(sv(row_raw, 1))
        l_leaf = parse_amount(sv(row_raw, left_leaf_col))
        l_grp  = parse_amount(sv(row_raw, left_grp_col))

        if l0 and not _is_skip(l0):
            if l_grp != 0.0 and l_leaf == 0.0:
                left_group = l0
            elif l_leaf != 0.0:
                left_group = l0

        if l1 and not _is_skip(l1) and l_leaf != 0.0:
            e = TrialBalanceEntry(account_name=l1, group=left_group)
            if is_pl:
                e.debit  = abs(l_leaf) if l_leaf >= 0 else 0.0
                e.credit = abs(l_leaf) if l_leaf <  0 else 0.0
            else:
                e.credit = l_leaf if l_leaf >= 0 else 0.0
                e.debit  = abs(l_leaf) if l_leaf < 0 else 0.0
            entries.append(e)

        r5 = _clean(sv(row_raw, 5))
        r6 = _clean(sv(row_raw, 6))
        r_leaf = parse_amount(sv(row_raw, right_leaf_col))
        r_grp  = parse_amount(sv(row_raw, right_grp_col))

        if r5 and not _is_skip(r5):
            if r_grp != 0.0 and r_leaf == 0.0:
                right_group = r5
            elif r_leaf != 0.0:
                right_group = r5

        if r6 and not _is_skip(r6) and r_leaf != 0.0:
            e = TrialBalanceEntry(account_name=r6, group=right_group)
            if is_pl:
                e.credit = r_leaf if r_leaf >= 0 else 0.0
                e.debit  = abs(r_leaf) if r_leaf < 0 else 0.0
            else:
                e.debit  = r_leaf if r_leaf >= 0 else 0.0
                e.credit = abs(r_leaf) if r_leaf < 0 else 0.0
            entries.append(e)

    return entries


# ── Format D: PDF two-column cell parsing ─────────────────────────────────────

def _parse_name_amount_lines(text: str) -> List[Tuple[str, float]]:
    _skip = {'nan', 'particulars', 'amount rs.', 'amount rs', 'amount', '',
             'total', 'grand total'}
    results = []
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        m = _NAME_AMT_RE.match(line)
        if m:
            name    = _clean(m.group(1).strip())
            amt_str = m.group(2).strip()
            if amt_str.startswith('(') and amt_str.endswith(')'):
                amt_str = '-' + amt_str[1:-1]
            amt = parse_amount(amt_str)
            if name and name.lower() not in _skip:
                results.append((name, amt))
    return results


def _detect_pdf_groups(pairs: List[Tuple[str, float]]) -> List[Tuple[str, float, str]]:
    n = len(pairs)
    is_header = [False] * n

    for i in range(n):
        name, amt = pairs[i]
        abs_amt = abs(amt)
        if abs_amt < 0.01:
            continue
        running_sign = running_abs = 0.0
        for j in range(i + 1, min(i + 500, n)):
            running_sign += pairs[j][1]
            running_abs  += abs(pairs[j][1])
            if abs(running_sign - amt) < 0.50 or abs(running_abs - abs_amt) < 0.50:
                is_header[i] = True
                break
            if running_abs > abs_amt * 4 + 1000:
                break

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
    all_left:  List[Tuple[str, float]] = []
    all_right: List[Tuple[str, float]] = []

    for row in table:
        if not row or len(row) < 2:
            continue
        left_cell  = str(row[0] or '')
        right_cell = str(row[1] or '')
        if 'particulars' in left_cell.lower().strip() and len(left_cell) < 60:
            continue
        all_left.extend(_parse_name_amount_lines(left_cell))
        all_right.extend(_parse_name_amount_lines(right_cell))

    entries: List[TrialBalanceEntry] = []
    for name, amt, grp in _detect_pdf_groups(all_left):
        e = TrialBalanceEntry(account_name=name, group=grp)
        e.credit = amt if amt > 0 else 0.0
        e.debit  = abs(amt) if amt < 0 else 0.0
        entries.append(e)
    for name, amt, grp in _detect_pdf_groups(all_right):
        e = TrialBalanceEntry(account_name=name, group=grp)
        e.debit  = amt if amt > 0 else 0.0
        e.credit = abs(amt) if amt < 0 else 0.0
        entries.append(e)
    return entries



# ── Format E: Tally Two-Column P&L / Balance Sheet (pdfplumber word bbox) ───
# Uses x-position of amount words to identify group-level totals:
#   Left side  (x < mid ~305): group totals have amount x >= 220 → Debit (Expenses)
#   Right side (x >= mid ~305): group totals have amount x >= 490 → Credit (Income/Sales)

def _detect_tally_two_col(pdf_path: str) -> bool:
    """Return True if the PDF looks like a Tally two-column P&L / BS."""
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return False
            text = (pdf.pages[0].extract_text() or '').lower()
            return ('profit' in text and 'loss' in text) or \
                   ('liabilit' in text and 'asset' in text) or \
                   ('particulars' in text and any(c in text for c in ['purchases', 'sales', 'opening stock']))
    except Exception:
        return False


def _parse_tally_two_col_pdf(pdf_path: str) -> List[TrialBalanceEntry]:
    """
    Parse Tally-exported two-column P&L or Balance Sheet PDF.
    Uses word bounding box x-positions to identify:
      - Left side (x < ~51% page width) = Debit/Expense side
      - Right side (x >= ~51% page width) = Credit/Income side
      - Group totals = amounts at the outermost x column per side
        Left: amount x >= 220 (sub-items have x ~160-200)
        Right: amount x >= 490 (sub-items have x ~430-480)
    """
    from collections import defaultdict as _dd
    entries: List[TrialBalanceEntry] = []

    _SKIP_KW = [
        'particulars', 'total', 'gross profit', 'nett profit', 'net profit',
        'e-mail', 'radha', 'govind', 'synthetics', 'ichalkaranji', 'bohara',
        '2/27', 'profit & loss', 'balance sheet', 'profit and loss',
    ]

    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                mid_x  = page.width * 0.51
                left_min_amt_x  = page.width * 0.37   # ~220 on 595-wide page
                right_min_amt_x = page.width * 0.82   # ~490 on 595-wide page

                words = page.extract_words(keep_blank_chars=False)
                rows: dict = _dd(list)
                for w in words:
                    y_key = round(float(w["top"]) / 4) * 4
                    rows[y_key].append(w)

                for y in sorted(rows.keys()):
                    row_w = sorted(rows[y], key=lambda w: float(w["x0"]))
                    left_w  = [w for w in row_w if float(w["x0"]) < mid_x]
                    right_w = [w for w in row_w if float(w["x0"]) >= mid_x]

                    for ws, is_right in [(left_w, False), (right_w, True)]:
                        if not ws:
                            continue
                        min_amt_x = right_min_amt_x if is_right else left_min_amt_x

                        # Find last numeric word at or beyond min_amt_x
                        amt_word = None
                        for w in reversed(ws):
                            txt = w["text"].replace("(-)", "")
                            val = parse_amount(txt)
                            if val > 0 and float(w["x0"]) >= min_amt_x:
                                amt_word = w
                                break
                        if not amt_word:
                            continue

                        neg = "(-)" in amt_word["text"]
                        amt = parse_amount(amt_word["text"].replace("(-)", ""))
                        name = " ".join(
                            w["text"] for w in ws
                            if float(w["x0"]) < float(amt_word["x0"])
                        ).strip()

                        if not name or len(name) < 2:
                            continue
                        nl = name.lower()
                        if any(kw in nl for kw in _SKIP_KW):
                            continue
                        if re.match(r"^\d{1,2}-[A-Za-z]", name):
                            continue

                        final_amt = -amt if neg else amt
                        e = TrialBalanceEntry(account_name=name, group="")
                        if is_right:
                            # Right side = Credit (Income/Sales/Assets)
                            if final_amt < 0:
                                e.debit  = abs(final_amt)
                            else:
                                e.credit = final_amt
                        else:
                            # Left side = Debit (Expenses/Liabilities/Purchases)
                            if final_amt < 0:
                                e.credit = abs(final_amt)
                            else:
                                e.debit  = final_amt
                        entries.append(e)

    except Exception as ex:
        logger.warning("Tally two-col PDF parse error: %s", ex)

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
                df = pd.read_csv(file_path, header=None, dtype=str,
                                 engine='python', on_bad_lines='skip')
            else:
                raise UnsupportedFileTypeError(f"Unsupported file type: {ext}")
        except UnsupportedFileTypeError:
            raise
        except Exception as e:
            raise FileParseError(f"Failed to read file: {e}")

        df = df.fillna('').map(lambda x: str(x).strip())
        return self._parse_dataframe(df, file_path)

    def _parse_dataframe(self, df: pd.DataFrame, source: str = '') -> List[TrialBalanceEntry]:
        # Strategy A: Standard Dr/Cr header
        header_row, col_map = _detect_columns(df)
        if header_row >= 0:
            entries = self._extract_standard(df, header_row, col_map)
            if entries:
                logger.info("Parsed %d TB entries (standard) from %s", len(entries), source)
                return entries

        # Strategy B: Balance Sheet Liabilities/Assets layout
        if _is_balance_sheet_format(df):
            entries = _parse_balance_sheet_format(df)
            if entries:
                logger.info("Parsed %d TB entries (BS/PL) from %s", len(entries), source)
                return entries

        # Strategy C: PARTICULARS/AMOUNT two-column
        if _is_particulars_two_column(df):
            entries = _parse_particulars_two_column(df)
            if entries:
                logger.info("Parsed %d TB entries (PARTICULARS two-col) from %s", len(entries), source)
                return entries

        raise FileParseError(
            "Could not detect column headers (Debit/Credit). "
            "Supported formats:\n"
            "  • Standard TB: 'Amount (Dr)'/'Amount (Cr)' or 'Debit'/'Credit'\n"
            "  • Balance Sheet/P&L (Liabilities/Assets headers)\n"
            "  • Balance Sheet/P&L (PARTICULARS/AMOUNT two-column layout)"
        )

    def _extract_standard(
        self, df: pd.DataFrame, header_row: int, col_map: Dict[str, int]
    ) -> List[TrialBalanceEntry]:
        # Use merged-header version of df
        df2 = _merge_header_rows(df)
        entries = []
        current_group = ''

        for row_idx in range(header_row + 1, len(df2)):
            row = df2.iloc[row_idx]
            account_val = _clean(str(row.iloc[col_map['account']]))
            debit_val   = str(row.iloc[col_map['debit']]).strip()
            credit_val  = str(row.iloc[col_map['credit']]).strip()

            if not account_val and not debit_val and not credit_val:
                continue

            # Handle "Cr"/"Dr" suffix in amount columns
            debit_val  = re.sub(r'\s*Dr\s*$', '', debit_val,  flags=re.I)
            credit_val = re.sub(r'\s*Cr\s*$', '', credit_val, flags=re.I)

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

            for key in ('opening_dr', 'opening_cr', 'closing_dr', 'closing_cr'):
                if key in col_map:
                    setattr(e, key, parse_amount(str(row.iloc[col_map[key]])))

            entries.append(e)
        return entries

    # ── PDF ───────────────────────────────────────────────────────────────────

    def _parse_pdf(self, file_path: str) -> List[TrialBalanceEntry]:
        import pdfplumber

        # ── Try Format E: Tally two-column P&L/BS (word bbox method) first ──────
        if _detect_tally_two_col(file_path):
            e_entries = _parse_tally_two_col_pdf(file_path)
            if e_entries:
                logger.info("Parsed %d entries via Tally two-col (Format E)", len(e_entries))
                return e_entries

        all_entries: List[TrialBalanceEntry] = []
        two_col_left:  List[Tuple[str, float]] = []
        two_col_right: List[Tuple[str, float]] = []
        is_two_col = False
        std_entries: List[TrialBalanceEntry] = []

        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables() or []
                    for table in tables:
                        if not table or len(table) < 2:
                            continue

                        if len(table[0]) == 2:
                            hdr = ' '.join(str(c or '').lower() for c in table[0])
                            if 'particulars' in hdr or 'amount' in hdr:
                                is_two_col = True
                                for row in table:
                                    if not row or len(row) < 2:
                                        continue
                                    lc = str(row[0] or '')
                                    rc = str(row[1] or '')
                                    if 'particulars' in lc.lower().strip() and len(lc) < 60:
                                        continue
                                    two_col_left.extend(_parse_name_amount_lines(lc))
                                    two_col_right.extend(_parse_name_amount_lines(rc))
                                continue

                        for hdr_idx, row in enumerate(table[:5]):
                            if not row:
                                continue
                            hdr_low = [str(c or '').lower().strip() for c in row]
                            if (any(_col_has(c, _DEBIT_KEYWORDS)  for c in hdr_low) and
                                    any(_col_has(c, _CREDIT_KEYWORDS) for c in hdr_low)):
                                ents = self._extract_from_pdf_table(table, hdr_idx, hdr_low)
                                if ents:
                                    std_entries.extend(ents)
                                break
        except Exception as e:
            logger.warning("pdfplumber table extraction failed: %s", e)

        if is_two_col and (two_col_left or two_col_right):
            for name, amt, grp in _detect_pdf_groups(two_col_left):
                e = TrialBalanceEntry(account_name=name, group=grp)
                e.credit = amt if amt > 0 else 0.0
                e.debit  = abs(amt) if amt < 0 else 0.0
                all_entries.append(e)
            for name, amt, grp in _detect_pdf_groups(two_col_right):
                e = TrialBalanceEntry(account_name=name, group=grp)
                e.debit  = amt if amt > 0 else 0.0
                e.credit = abs(amt) if amt < 0 else 0.0
                all_entries.append(e)

        if std_entries:
            all_entries.extend(std_entries)

        if all_entries:
            logger.info("PDF parsed %d entries via table extraction", len(all_entries))
            return all_entries

        # Text-line fallback
        raw_lines = []
        for lib in ('pypdf', 'pdfplumber'):
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
        padded   = [r + [''] * (max_cols - len(r)) for r in raw_lines]
        df = pd.DataFrame(padded, dtype=str).fillna('').map(lambda x: str(x).strip())

        try:
            return self._parse_dataframe(df, file_path)
        except FileParseError:
            pass

        entries = self._parse_pdf_text_two_column(raw_lines)
        if entries:
            return entries

        raise FileParseError(
            "Could not detect Debit/Credit columns in the PDF. "
            "Ensure the file has standard headers like 'Debit'/'Credit', "
            "'Dr'/'Cr', 'Amount (Dr)'/'Amount (Cr)', or 'Liabilities'/'Assets'."
        )

    def _parse_pdf_text_two_column(
        self, raw_lines: List[List[str]]
    ) -> List[TrialBalanceEntry]:
        all_left:  List[Tuple[str, float]] = []
        all_right: List[Tuple[str, float]] = []

        for parts in raw_lines:
            if len(parts) < 2:
                continue
            lt = parts[0].strip()
            rt = parts[1].strip()
            if any(kw in lt.lower() for kw in ('particulars', 'amount rs', 'page no')):
                continue
            all_left.extend(_parse_name_amount_lines(lt))
            all_right.extend(_parse_name_amount_lines(rt))

        entries: List[TrialBalanceEntry] = []
        for name, amt, grp in _detect_pdf_groups(all_left):
            e = TrialBalanceEntry(account_name=name, group=grp)
            e.credit = amt if amt > 0 else 0.0
            e.debit  = abs(amt) if amt < 0 else 0.0
            entries.append(e)
        for name, amt, grp in _detect_pdf_groups(all_right):
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
            if debit_col  < 0 and _col_has(cell, _DEBIT_KEYWORDS):   debit_col  = i
            elif credit_col < 0 and _col_has(cell, _CREDIT_KEYWORDS): credit_col = i
            elif account_col < 0 and _col_has(cell, _ACCOUNT_KEYWORDS): account_col = i
        if account_col < 0: account_col = 0
        if debit_col  < 0 or credit_col < 0:
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
