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
    'print date', 'print date :', 'page no', 'page no.',
}
# Profit/Loss rows are skipped ONLY when they carry no amount (derived rows).
# When they DO have an amount (e.g. "Net Profit" Dr in a TB, "Profit & Loss A/c"
# Dr balance in Tally) they must be kept as real entries.
_PROFIT_ONLY_KW = {'gross profit', 'net profit', 'net loss'}

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


_TOTAL_KW  = {'total', 'grand total', 'sub total', 'subtotal', 'net total'}
_PROFIT_KW = {'gross profit', 'net profit', 'net loss'}

# Net-profit-like ledger names found on the expense side of a P&L:
#   "NET PROFIT", "NETT PROFIT", "Z. Net Profit A/c", "PROFIT & LOSS A/C",
#   "PROFIT AND LOSS AIC" (OCR/typo variants). Per the TB business rules these
# must be KEPT and grouped into CAPITAL (debit side).
_NET_PROFIT_RE   = re.compile(r'net{1,2}\s*profit|net\s*loss|profit\s*(?:&|and)\s*loss\s*a', re.I)
_GROSS_PROFIT_RE = re.compile(r'gross\s*profit', re.I)


def _collapse_side_pairs(
    pairs: List[Tuple[str, float]],
    flags: Optional[List[Optional[bool]]] = None,
) -> List[Tuple[str, float, str]]:
    """
    Given ONE SIDE of a two-sided statement as an ordered list of
    (name, amount) rows, detect group-header/subtotal rows — rows whose amount
    equals the sum of the immediately following run of rows — and remove them
    so they are not double counted. Header names become the group label of the
    rows they cover. Zero-amount named rows are treated as plain group labels.

    flags: optional per-row structural hint. True = header candidate (name in
    the outer column, not indented), False = leaf (inner column or indented),
    None = unknown. A row flagged False may NOT absorb a SINGLE following row
    of equal amount — that pattern is two sibling leaves that happen to share
    an amount (e.g. two GST input ledgers at 414.90), not a header+leaf pair.
    Multi-row sum matches are still allowed regardless of flag, since a
    multi-row coincidence is far less likely than an equal-amount pair.

    Returns [(name, amount, group), ...] containing only leaf rows.

    This is the structural fix for statements where group totals are printed
    in the SAME column as their detail rows (e.g. "PURCHASE A/C 182270.44"
    followed by "MILL STORE A/C 182270.44").
    """
    n = len(pairs)
    is_hdr = [False] * n

    def _flag(i):
        if flags is None or i >= len(flags):
            return None
        return flags[i]

    for i, (name, amt) in enumerate(pairs):
        if abs(amt) < 0.01:
            is_hdr[i] = True            # label-only row → group label
            continue
        run_sign = run_abs = 0.0
        run_rows = 0
        for j in range(i + 1, min(i + 500, n)):
            aj = pairs[j][1]
            if abs(aj) < 0.01:
                continue
            run_sign += aj
            run_abs  += abs(aj)
            run_rows += 1
            tol = max(1.0, abs(amt) * 0.002)
            if abs(run_sign - amt) <= tol or abs(run_abs - abs(amt)) <= tol:
                if run_rows == 1:
                    child_name = pairs[j][0]
                    # Single-child fold: allow only when there is clear evidence
                    # this row is a structural group header, not an equal-amount sibling.
                    # Evidence 1: flag=True  → outer column in indented layout
                    # Evidence 2: same name  → exact duplicate row (e.g. Profit & Loss A/C
                    #             printed twice); absorb to deduplicate.
                    # Evidence 3: known group name → structural group (Opening Stock, etc.)
                    # Evidence 4: flag=False → definitely a leaf; block absorption.
                    # Otherwise (flag=None, no group name): treat as siblings → block.
                    if _flag(i) is True:
                        pass            # outer-column header → allow
                    elif name.strip().lower() == child_name.strip().lower():
                        pass            # exact duplicate → deduplicate
                    elif _is_tally_std_group(name):
                        pass            # known structural group → allow
                    else:
                        continue        # insufficient evidence → siblings; keep both
                is_hdr[i] = True
                break
            if run_abs > abs(amt) * 1.5 + 1000:
                break

    out: List[Tuple[str, float, str]] = []
    grp = ''
    for i, (name, amt) in enumerate(pairs):
        if _is_skip(name):
            continue
        if is_hdr[i]:
            grp = name
            continue
        out.append((name, amt, grp))
    return out


def _normalize_profit_leaves(
    leaves: List[Tuple[str, float, str]]
) -> List[Tuple[str, float, str]]:
    """
    Apply P&L profit-row business rules to a collapsed leaf list:
      - Drop GROSS PROFIT rows (c/o and b/f appear on both sides and cancel;
        they are derived figures, never TB ledgers).
      - Keep NET PROFIT / NET LOSS / 'PROFIT & LOSS A/C' rows but rename to
        NET PROFIT (or NET LOSS) and group them under CAPITAL (Note_TB rule 6).
    """
    out: List[Tuple[str, float, str]] = []
    for name, amt, grp in leaves:
        nl = name.lower()
        if _GROSS_PROFIT_RE.search(nl):
            continue
        if _NET_PROFIT_RE.search(nl):
            label = 'NET LOSS' if 'loss' in nl and 'profit' not in nl else 'NET PROFIT'
            out.append((label, amt, 'CAPITAL'))
            continue
        out.append((name, amt, grp))
    return out


def _kw_match(nl: str, kws) -> bool:
    """Whole-word keyword match. Substring matching wrongly hits names like
    'VIVEK RAMKUMAR TOTALA' (contains 'total') or 'Packing Matrial' ('trial')."""
    for kw in kws:
        if re.search(r'(?<![a-z])' + re.escape(kw) + r'(?![a-z])', nl):
            return True
    return False


def _is_skip(name: str) -> bool:
    """Return True only for pure total rows. Profit/Loss rows are NOT skipped
    here — callers must check _is_profit_row separately and skip only when amount=0."""
    nl = name.lower().strip()
    return _kw_match(nl, _SKIP_ROW_KW)


def _is_total_row(name: str) -> bool:
    nl = name.lower().strip()
    return _kw_match(nl, _TOTAL_KW)


def _is_profit_row(name: str) -> bool:
    nl = name.lower().strip()
    return _kw_match(nl, _PROFIT_KW)


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
                # ── Refine amount columns by BODY DATA ────────────────────────
                # Header labels sometimes sit one column left of where the
                # values actually land (merged cells / export quirks), which
                # silently drops an entire side (e.g. 'Credit' header at col 9
                # but every credit value in col 10). For each amount column,
                # if it has (almost) no numeric body values but a nearby
                # column to the right has many, shift to the data column.
                def _numeric_count(ci):
                    if ci < 0 or ci >= df2.shape[1]:
                        return 0
                    ct = 0
                    for ri2 in range(row_idx + 1, len(df2)):
                        if _is_numeric(str(df2.iloc[ri2, ci]).strip()):
                            ct += 1
                    return ct

                taken = {col_map.get('account', 0)}
                for key in ('debit', 'credit'):
                    cur = col_map[key]
                    other = col_map['credit'] if key == 'debit' else col_map['debit']
                    best, best_ct = cur, _numeric_count(cur)
                    for cand in (cur + 1, cur + 2):
                        if cand == other or cand in taken or cand >= df2.shape[1]:
                            continue
                        ct = _numeric_count(cand)
                        if ct > best_ct * 2 + 2:
                            best, best_ct = cand, ct
                    if best != cur:
                        logger.info("Column refine: '%s' header at col %d but data "
                                    "found at col %d — using col %d", key, cur, best, best)
                        col_map[key] = best
                    taken.add(col_map[key])
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

    # Group-total fallback bookkeeping:
    # When a group header row carries its total (sub=0, total≠0) but every
    # detail row under it has no parseable amount in the numeric columns
    # (e.g. "YARN STOCK :-42518020" — figure embedded in the NAME text),
    # the whole group would silently vanish. Track each group's header
    # total and how many entries it actually produced; emit the header
    # total itself as a single entry for any group that produced none.
    # NOTE: we deliberately do NOT parse amounts embedded in names — in
    # Tally P&L exports those are often CLOSING values printed beside
    # OPENING rows, so the header total is the only trustworthy figure.
    left_group_totals:  Dict[str, float] = {}
    right_group_totals: Dict[str, float] = {}
    left_group_counts:  Dict[str, int]   = {}
    right_group_counts: Dict[str, int]   = {}

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
                left_group_totals[left_group] = left_total
                left_group_counts.setdefault(left_group, 0)
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
                if left_group:
                    left_group_counts[left_group] = left_group_counts.get(left_group, 0) + 1

        right_name  = _clean(sv(row_raw, rn_col))
        right_sub   = parse_amount(sv(row_raw, rs_col))
        right_total = parse_amount(sv(row_raw, rt_col))

        if right_name and right_name.lower() not in ('nan', 'asset', 'assets', 'amount', ''):
            if _is_skip(right_name):
                pass
            elif right_sub == 0.0 and right_total != 0.0:
                right_group = right_name
                right_group_totals[right_group] = right_total
                right_group_counts.setdefault(right_group, 0)
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
                if right_group:
                    right_group_counts[right_group] = right_group_counts.get(right_group, 0) + 1

    # ── Group-total fallback ──────────────────────────────────────────────
    # Any group whose header carried a total but whose detail rows yielded
    # ZERO entries gets emitted once, at the header total, on the correct
    # side. Fixes e.g. OPENING STOCK ₹5,09,31,678.67 being dropped because
    # YARN/CLOTH detail amounts live inside the name text, not the columns.
    for grp, total in left_group_totals.items():
        if left_group_counts.get(grp, 0) == 0 and total != 0.0:
            e = TrialBalanceEntry(account_name=grp, group=grp)
            if is_pl:
                e.debit  = abs(total) if total >= 0 else 0.0
                e.credit = abs(total) if total < 0  else 0.0
            else:
                e.credit = total if total >= 0 else 0.0
                e.debit  = abs(total) if total < 0  else 0.0
            entries.append(e)
            logger.info("Group-total fallback (left/%s): %s = %.2f",
                        'Dr' if is_pl else 'Cr', grp, total)

    for grp, total in right_group_totals.items():
        if right_group_counts.get(grp, 0) == 0 and total != 0.0:
            e = TrialBalanceEntry(account_name=grp, group=grp)
            if is_pl:
                e.credit = total if total >= 0 else 0.0
                e.debit  = abs(total) if total < 0  else 0.0
            else:
                e.debit  = total if total >= 0 else 0.0
                e.credit = abs(total) if total < 0  else 0.0
            entries.append(e)
            logger.info("Group-total fallback (right/%s): %s = %.2f",
                        'Cr' if is_pl else 'Dr', grp, total)

    return entries


# ── Format C: PARTICULARS/AMOUNT two-column XLS ───────────────────────────────

def _is_particulars_two_column(df: pd.DataFrame) -> bool:
    row_idx, ok = _find_particulars_anchors(df)
    return ok and row_idx >= 0


_NAME_HEADERS = ('particular', 'account name', 'account head', 'ledger',
                 'name of account', 'account')


def _find_particulars_header_row(df: pd.DataFrame) -> int:
    row_idx, _ = _find_particulars_anchors(df)
    return row_idx


def _find_particulars_anchors(df: pd.DataFrame):
    """Locate a two-sided (Liabilities|Amount || Assets|Amount) header row.

    Matches PARTICULARS / Account Name / Ledger style headers with two or more
    Amount columns. Returns (row_idx, True) or (-1, False).
    """
    for row_idx in range(min(20, len(df))):
        cells = [str(c).lower().strip() for c in df.iloc[row_idx]]
        n_name = sum(1 for c in cells if any(h in c for h in _NAME_HEADERS))
        n_amt  = sum(1 for c in cells if 'amount' in c)
        if (n_name >= 1 and n_amt >= 2) or (n_name >= 2 and n_amt >= 1):
            return row_idx, True
    return -1, False


def _filename_is_pl(source: str):
    """Filename hint: P&L vs Balance Sheet. Returns True (P&L), False (BS) or None.

    Per the client rule: a P&L file means Left=Dr / Right=Cr; a Balance Sheet
    means Left=Cr / Right=Dr. Filenames carry 'pl'/'p&l'/'profit'/'trading' or
    'bs'/'bl'/'balance'.

    BL/bl = Balance (Sheet) — treated same as BS.
    """
    s = (source or '').lower()
    pl_keys = ('p & l', 'p&l', 'profit', 'pandl', 'p and l', '-pl', '_pl',
               ' pl', 'pl.', 'p.l', 'p_l', 'trading', 'p-l')
    bs_keys = ('balance', '-bs', '_bs', ' bs', 'bs.', 'b.s', 'b_s',
               '-bl', '_bl', ' bl', 'bl.', 'b.l', 'b_l', 'b/s', 'bsheet')
    if any(k in s for k in pl_keys):
        return True
    if any(k in s for k in bs_keys):
        return False
    return None


def _parse_particulars_two_column(df: pd.DataFrame, is_pl_hint=None) -> List[TrialBalanceEntry]:
    hdr_row, ok = _find_particulars_anchors(df)
    if not ok or hdr_row < 0:
        return []

    # Filename hint (client rule) takes precedence over content heuristics.
    is_pl = is_pl_hint if is_pl_hint is not None else _is_pl_document(df, hdr_row)
    ncols = df.shape[1]

    # ── Infer column roles EMPIRICALLY from the body (header label positions
    #    are unreliable — some templates compress headers while data spreads
    #    across other columns). Classify each column as Text or Numeric. ──────
    text_ct = [0] * ncols
    num_ct  = [0] * ncols
    for ri in range(hdr_row + 1, len(df)):
        row = list(df.iloc[ri])
        for c in range(min(ncols, len(row))):
            v = str(row[c]).strip()
            if not v or v.lower() == 'nan':
                continue
            if _is_numeric(row[c]) and parse_amount(v) != 0.0:
                num_ct[c] += 1
            elif not _is_numeric(row[c]) and len(v) >= 2:
                text_ct[c] += 1

    # Cluster consecutive same-role columns (small gaps tolerated).
    clusters = []  # list of (role, [cols])
    for c in range(ncols):
        if num_ct[c] == 0 and text_ct[c] == 0:
            continue
        role = 'N' if num_ct[c] >= text_ct[c] else 'T'
        if clusters and clusters[-1][0] == role and c - clusters[-1][1][-1] <= 2:
            clusters[-1][1].append(c)
        else:
            clusters.append((role, [c]))

    name_clusters = [cols for role, cols in clusters if role == 'T']
    amt_clusters  = [cols for role, cols in clusters if role == 'N']
    if not name_clusters or not amt_clusters:
        return []

    # ── Match name clusters to amount clusters by proximity ───────────────────
    # Filter noise clusters: too few entries, or only 'TOTAL'/'GRAND TOTAL' text.
    _NOISE_TEXTS = {'total', 'grand total', 'sr no', 'sr.no', 'page no', 'page',
                    'print date', 'print'}

    def _is_noise_cluster(cols):
        total_texts = sum(text_ct[c] for c in cols)
        if total_texts <= 1:
            return True
        # Scan full body for text values in these columns
        sample = []
        for ri in range(hdr_row + 1, len(df)):
            for c in cols:
                v = str(df.iloc[ri, c]).strip()
                if v and v.lower() not in ('nan', ''):
                    sample.append(v.lower())
        # Empty sample (no values found in body) → noise
        if not sample:
            return True
        # All values are noise keywords → noise
        if all(any(n in s for n in _NOISE_TEXTS) for s in sample):
            return True
        return False

    def _is_noise_amt_cluster(cols):
        """Skip amount clusters that are clearly TOTAL/summary-only columns."""
        total_nums = sum(num_ct[c] for c in cols)
        return total_nums <= 2  # only a TOTAL row — not a real data column

    valid_name_clusters = [cols for cols in name_clusters if not _is_noise_cluster(cols)]
    valid_amt_clusters  = [cols for cols in amt_clusters  if not _is_noise_amt_cluster(cols)]
    if not valid_name_clusters:
        valid_name_clusters = name_clusters
    if not valid_amt_clusters:
        valid_amt_clusters = amt_clusters

    # Pair each valid name cluster with the nearest SUBSTANTIAL amt cluster to its right.
    # If no amt cluster is to the right, fall back to the nearest at all.
    sides = []  # (name_cols, amt_cols, positive_to_debit, indented)

    def _indented(name_cols):
        return sum(1 for c in name_cols if text_ct[c] > 0) > 1

    used_amt_idx = set()
    for name_cols in valid_name_clusters[:2]:   # max 2 sides
        name_max = max(name_cols)
        # Prefer amt clusters starting AFTER this name cluster
        candidates = [(i, cols) for i, cols in enumerate(valid_amt_clusters)
                      if min(cols) > name_max and i not in used_amt_idx]
        if not candidates:
            # Fallback: nearest available
            candidates = [(i, cols) for i, cols in enumerate(valid_amt_clusters)
                          if i not in used_amt_idx]
        if not candidates:
            continue
        best_i, best_cols = min(candidates, key=lambda x: abs(min(x[1]) - name_max))
        used_amt_idx.add(best_i)
        is_dr = is_pl if len(sides) == 0 else not is_pl
        sides.append((name_cols, best_cols, is_dr, _indented(name_cols)))

    # ── Collect ordered (name, amount) rows PER SIDE, then post-process ──────
    # Many statements print group totals in the SAME column as detail rows
    # (e.g. "PURCHASE A/C 182270.44" followed by "MILL STORE A/C 182270.44"),
    # so structural emit-as-you-go double counts. Instead: collect each side's
    # row sequence, collapse subtotal/header rows by sum-matching, then apply
    # the profit-row business rules, and only then emit entries.
    # Build per-side lists AND flags (True=outer=header candidate, False=inner=leaf)
    side_pairs: List[List[Tuple[str, float]]] = [[] for _ in sides]
    side_flags: List[List[Optional[bool]]] = [[] for _ in sides]

    def _rightmost_amt(row, cols):
        val = 0.0
        for c in cols:
            if 0 <= c < len(row):
                a = parse_amount(str(row[c]).strip())
                if a != 0.0:
                    val = a
        return val

    for row_idx in range(hdr_row + 1, len(df)):
        row = list(df.iloc[row_idx])
        for si, (name_cols, amt_cols, ptd, indented) in enumerate(sides):
            present = []
            for c in name_cols:
                if 0 <= c < len(row):
                    v = _clean(str(row[c]).strip())
                    if v and v.lower() != 'nan' and not _is_numeric(row[c]):
                        present.append((c, v))
            amt = _rightmost_amt(row, amt_cols)
            if not present:
                continue

            is_outer_only = False       # flag: text only in outer (group) col
            if indented:
                outer = min(name_cols)
                group_txt = next((t for c, t in present if c == outer), None)
                leaf_txt  = next((t for c, t in present if c > outer), None)
                # Some exports DUPLICATE the name into the next column,
                # truncating the copy from the left ("TOTAL :" → "L :",
                # "GROSS PROFIT" → "S PROFIT"). When one text is a tail of
                # the other they are the SAME name, not group+leaf — keep
                # the longer (full) one so total/profit filters can match.
                if group_txt and leaf_txt:
                    g_n = re.sub(r'\s+', ' ', group_txt).lower()
                    l_n = re.sub(r'\s+', ' ', leaf_txt).lower()
                    if g_n.endswith(l_n) or l_n.endswith(g_n) \
                            or g_n in l_n or l_n in g_n:
                        name = group_txt if len(group_txt) >= len(leaf_txt) \
                            else leaf_txt
                        is_outer_only = True   # it's ONLY in outer col logically
                    else:
                        name = leaf_txt
                else:
                    if group_txt and not leaf_txt:
                        is_outer_only = True   # actual group-header row
                    name = leaf_txt or group_txt
            else:
                name = present[0][1]

            if not name or _is_skip(name):
                continue
            # Skip single-char noise like '.' used as total separator rows
            if len(name.strip('.').strip()) == 0:
                continue
            # Skip print/page footer rows
            if re.match(r'^(print\s+date|page\s+no|page\s*:|end\s+of\s+report)', name.lower()):
                continue
            # Column-header echoes act as group-label resets
            if name.lower() in ('particulars', 'amount', 'amount rs', 'amount rs.'):
                side_pairs[si].append(('', 0.0))
                side_flags[si].append(True)
                continue
            # flag: True=outer/header row, False=inner/leaf, None=unknown
            flag = True if is_outer_only else (False if indented else None)
            side_pairs[si].append((name, amt))
            side_flags[si].append(flag)

    entries: List[TrialBalanceEntry] = []
    for si, (name_cols, amt_cols, ptd, indented) in enumerate(sides):
        leaves = _collapse_side_pairs(side_pairs[si], flags=side_flags[si])
        leaves = _normalize_profit_leaves(leaves)
        for name, amt, grp in leaves:
            if not name:
                continue
            e = TrialBalanceEntry(account_name=name, group=grp)
            if ptd:
                e.debit, e.credit = (abs(amt), 0.0) if amt >= 0 else (0.0, abs(amt))
            else:
                e.credit, e.debit = (abs(amt), 0.0) if amt >= 0 else (0.0, abs(amt))
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

    # Strong additional signal: a line whose name is itself a canonical
    # Trading/P&L/Balance-Sheet SECTION is a group header (indentation is lost
    # in PDF text, so name matching is the reliable cue).
    _SECTIONS = {
        'opening stock', 'closing stock', 'purchase', 'purchases', 'purchase a/c',
        'sales', 'sales a/c', 'sale', 'direct expenses', 'indirect expenses',
        'manufacturing expenses', 'direct income', 'indirect income',
        'direct incomes', 'indirect incomes', 'trading account', 'capital account',
        'current assets', 'current liabilities', 'fixed assets', 'investments',
        'loans & advances', 'loans and advances', 'sundry debtors',
        'sundry creditors', 'unsecured loans', 'secured loans', 'reserves & surplus',
    }
    for i, (name, amt) in enumerate(pairs):
        if name.strip().lower() in _SECTIONS:
            is_header[i] = True

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



# ── Format H: PARTICULARS/AMOUNT two-sided PDF, indentation-based ────────────
# (Genius / Miracle style BS & P&L exports, e.g. UMA TEXCOM bs.pdf / pl.pdf)
#
# Layout facts this parser relies on (verified by word x-positions):
#   • Header row: "PARTICULARS  AMOUNT Rs.  PARTICULARS  AMOUNT RS."
#   • Group rows and detail rows share the SAME right-aligned amount column,
#     so amount x-position CANNOT distinguish them (Format E double-counts).
#   • The real signal is NAME INDENTATION: group names start at the side's
#     left margin; detail names are indented a few points to the right.
#   • Left side semantics:  BS → Liabilities (Cr) | P&L → Expenses (Dr)
#     Right side semantics: BS → Assets (Dr)      | P&L → Income (Cr)
#   • Parenthesised amounts are negative and flip the column.
#
# Only DETAIL rows become entries (group = enclosing header). A group whose
# details never parsed falls back to the group-header total — same safety
# net as Format B. Derived rows (TOTAL / GROSS PROFIT / NET PROFIT) are
# skipped; the service re-derives net profit and plugs it per TB rules.

_H_SKIP_KW = ('particulars', 'total', 'gross profit', 'nett profit',
              'net profit', 'net loss', 'print date', 'page no', 'amount')


def _h_doc_is_pl(text: str) -> bool:
    t = text.lower()
    return ('profit' in t and 'loss' in t) or 'trading' in t


def _detect_particulars_two_col(pdf_path: str) -> bool:
    """True when page 1 has a PARTICULARS/AMOUNT header (1 or 2 sided)."""
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return False
            text = (pdf.pages[0].extract_text() or '').lower()
            return text.count('particulars') >= 1 and 'amount' in text
    except Exception:
        return False


def _parse_particulars_two_col_pdf(pdf_path: str) -> List[TrialBalanceEntry]:
    """
    Two-pass parse:
      Pass 1 — collect every (side, name_x0, name, amount) row across ALL
               pages, using tolerance-based row clustering (fixed y-buckets
               split wrapped rows whose name and amount baselines differ).
      Pass 2 — classify group vs detail using the DOCUMENT-GLOBAL minimum
               name x0 per side. Per-page minima misfire on continuation
               pages that contain only detail rows (everything gets
               promoted to group level and the grouping cascades).
    Groups persist across pages; only detail rows become entries; a group
    whose details never parse falls back to its header total.
    """
    import pdfplumber

    raw_rows = []          # (order, side, x0, name, amt)
    is_pl = False
    order = 0

    with pdfplumber.open(pdf_path) as pdf:
        if pdf.pages:
            is_pl = _h_doc_is_pl(pdf.pages[0].extract_text() or '')
        last_split_x = None

        for page in pdf.pages:
            words = page.extract_words(keep_blank_chars=False)
            if not words:
                continue

            # ── tolerance-based row clustering (gap > 3.5pt starts new row) ──
            words.sort(key=lambda w: (float(w['top']), float(w['x0'])))
            lines, cur, cur_top = [], [], None
            for w in words:
                t = float(w['top'])
                if cur_top is None or t - cur_top <= 3.5:
                    cur.append(w)
                    cur_top = t if cur_top is None else max(cur_top, t)
                else:
                    lines.append(cur)
                    cur, cur_top = [w], t
            if cur:
                lines.append(cur)

            # ── locate header line & side boundary ────────────────────────
            header_i, right_start = None, None
            for i, ln in enumerate(lines):
                ws = sorted(ln, key=lambda w: float(w['x0']))
                partis = [w for w in ws
                          if w['text'].strip().lower().startswith('particular')]
                if partis and any('amount' in w['text'].lower() for w in ws):
                    header_i = i
                    if len(partis) >= 2:
                        right_start = float(partis[1]['x0'])
                    break
            if header_i is None:
                if last_split_x is None:
                    continue          # no header seen yet in document
                split_x, header_i = last_split_x, -1
            else:
                if right_start is None:
                    right_start = page.width * 0.52
                split_x = right_start - 60
                last_split_x = split_x

            # ── extract (side, x0, name, amount) per line ──────────────────
            for ln in lines[header_i + 1:]:
                ws = sorted(ln, key=lambda w: float(w['x0']))
                for side, sws in (
                        ('L', [w for w in ws if float(w['x0']) < split_x]),
                        ('R', [w for w in ws if float(w['x0']) >= split_x])):
                    if not sws:
                        continue
                    last = sws[-1]
                    amt = parse_amount(last['text'])
                    if amt == 0.0 and not _PAREN_NUM_RE.match(last['text']):
                        continue                      # no trailing amount
                    name = _clean(' '.join(w['text'] for w in sws[:-1]).strip())
                    if not name or len(name) < 2:
                        continue
                    nl = name.lower()
                    if any(kw in nl for kw in _H_SKIP_KW):
                        continue
                    raw_rows.append((order, side, float(sws[0]['x0']), name, amt))
                    order += 1

    if not raw_rows:
        return []

    # ── document-global indent level per side ─────────────────────────────
    min_x = {}
    for _, side, x0, _, _ in raw_rows:
        min_x[side] = min(min_x.get(side, 1e9), x0)

    entries: List[TrialBalanceEntry] = []
    group_totals: Dict[Tuple[str, str], float] = {}
    group_counts: Dict[Tuple[str, str], int]   = {}
    cur_group = {'L': '', 'R': ''}

    for _, side, x0, name, amt in sorted(raw_rows):
        if (x0 - min_x[side]) < 3.0:          # at the side's left margin → group
            cur_group[side] = name
            key = (side, name)
            group_totals[key] = group_totals.get(key, 0.0) + amt
            group_counts.setdefault(key, 0)
            continue
        entries.append(_h_make_entry(name, cur_group[side], amt, side, is_pl))
        if cur_group[side]:
            gkey = (side, cur_group[side])
            group_counts[gkey] = group_counts.get(gkey, 0) + 1

    # ── group-total fallback: groups whose details never parsed ───────────
    for (side, gname), total in group_totals.items():
        if group_counts.get((side, gname), 0) == 0 and total != 0.0:
            entries.append(_h_make_entry(gname, gname, total, side, is_pl))
            logger.info("Format H group-total fallback (%s): %s = %.2f",
                        side, gname, total)
    return entries


_PAREN_NUM_RE = re.compile(r'^\([\d,]+\.?\d*\)$')


def _h_make_entry(name: str, group: str, amt: float,
                  side: str, is_pl: bool) -> TrialBalanceEntry:
    e = TrialBalanceEntry(account_name=name, group=group)
    # side semantics; negative amounts flip the column
    if is_pl:
        dr_side = (side == 'L')        # P&L: left = expenses (Dr)
    else:
        dr_side = (side == 'R')        # BS:  right = assets (Dr)
    if amt < 0:
        dr_side = not dr_side
        amt = abs(amt)
    if dr_side:
        e.debit = amt
    else:
        e.credit = amt
    return e


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



# ── Format F: Flat single-column "Name - Debit/Credit Amount" ──────────────────
# Handles any text-based list format regardless of PDF layout:
#   "Cash in Hand - Debit 50,000"
#   "Sales - Credit 800,000"
#   "Capital Account   Credit   950,000"
#   "SBI Bank 150,000 Dr"

_FLAT_PATTERNS = [
    # "Name - Debit 50,000" or "Name - Credit 800,000"
    re.compile(r'^(.+?)\s*[-–]\s*(debit|credit)\s+([\d,]+\.?\d*)$', re.IGNORECASE),
    # "Name  Debit  50,000" (2+ spaces)
    re.compile(r'^(.+?)\s{2,}(debit|credit)\s+([\d,]+\.?\d*)$', re.IGNORECASE),
    # "Name 50,000 Dr"
    re.compile(r'^(.+?)\s+([\d,]+\.?\d*)\s*(dr|cr)\.?$', re.IGNORECASE),
    # "Name Dr 50,000"
    re.compile(r'^(.+?)\s+(dr|cr)\.?\s+([\d,]+\.?\d*)$', re.IGNORECASE),
    # "Name (Dr) 50,000"
    re.compile(r'^(.+?)\s+\((dr|cr)\)\s+([\d,]+\.?\d*)$', re.IGNORECASE),
]

_FLAT_SKIP = re.compile(
    r'^(sample|test|trial balance|balance sheet|profit|loss|particulars|'
    r'sr\s*no|sl\s*no|s\.no|account\s*name|ledger|opening|closing|'
    r'total|grand total|prepared|date|for the|statement)\b',
    re.IGNORECASE,
)


def _is_flat_line_format(text: str) -> bool:
    """Return True if text contains flat Debit/Credit line entries."""
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    matches = 0
    for line in lines[:30]:
        for pat in _FLAT_PATTERNS:
            if pat.match(line):
                matches += 1
                break
    return matches >= 3


def _parse_flat_line_text(text: str) -> List[TrialBalanceEntry]:
    """Parse flat "Name - Debit/Credit Amount" text into TrialBalanceEntry list."""
    entries = []
    for line in text.split('\n'):
        line = line.strip()
        if not line or len(line) < 5:
            continue
        if _FLAT_SKIP.match(line):
            continue

        for i, pat in enumerate(_FLAT_PATTERNS):
            m = pat.match(line)
            if not m:
                continue

            if i in (0, 1):
                name      = m.group(1).strip().rstrip('-').strip()
                dc        = m.group(2).lower()
                amt_str   = m.group(3)
                direction = 'dr' if dc.startswith('d') else 'cr'
                amount    = parse_amount(amt_str)
            elif i == 2:
                name      = m.group(1).strip()
                amount    = parse_amount(m.group(2))
                direction = 'dr' if m.group(3).lower().startswith('d') else 'cr'
            else:  # i in (3, 4)
                name      = m.group(1).strip()
                direction = 'dr' if m.group(2).lower().startswith('d') else 'cr'
                amount    = parse_amount(m.group(3))

            if not name or amount <= 0:
                break
            if _FLAT_SKIP.match(name):
                break
            if len(name) < 2:
                break

            e = TrialBalanceEntry(account_name=name, group='')
            if direction == 'dr':
                e.debit  = amount
            else:
                e.credit = amount
            entries.append(e)
            break

    return entries


# ── Format G: Section-header "Name : Amount" (BS/PL extract without table) ───
# Handles PDFs/text with:
#   LIABILITIES
#   Capital Account : 3,00,000
#   ASSETS
#   Furniture : 2,00,000
#   PROFIT & LOSS ACCOUNT
#   Sales : 12,00,000

_SECTION_HEADERS = {
    # header keyword → (side, is_pl)
    'liabilit':  ('credit', False),
    'capital':   ('credit', False),
    'asset':     ('debit',  False),
    'income':    ('credit', True),
    'sales':     ('credit', True),
    'revenue':   ('credit', True),
    'receipt':   ('credit', True),
    'expenditure': ('debit', True),
    'expense':   ('debit',  True),
    'purchase':  ('debit',  True),
    'profit':    ('debit',  True),   # heading line, not entry
    'loss':      ('debit',  True),
}

_COLON_ENTRY_RE = re.compile(
    r'^(.+?)\s*[:\-]\s*([(\-]?[\d,]+\.?\d*[)]?)\s*$'
)


def _is_section_header_format(text: str) -> bool:
    """Return True if text has 'Name : Amount' lines under section headers."""
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    colon_hits = sum(1 for l in lines if _COLON_ENTRY_RE.match(l) and
                     any(c.isdigit() for c in l))
    # Also check that there's at least one recognisable section keyword
    has_section = any(
        any(kw in l.lower() for kw in _SECTION_HEADERS)
        for l in lines[:20]
    )
    return colon_hits >= 3 and has_section


def _parse_section_header_format(text: str) -> List[TrialBalanceEntry]:
    """
    Parse 'Section Header\\nName : Amount' text into TrialBalanceEntry list.

    Rules:
      LIABILITIES section  → credit side
      ASSETS section       → debit side
      PROFIT & LOSS ACCOUNT (income-like header) → Sales/Income → credit; Purchases/Expenses → debit
    """
    entries: List[TrialBalanceEntry] = []
    current_side  = 'debit'   # default
    current_group = ''
    in_pl         = False

    _CREDIT_SECTIONS = {'liabilit', 'capital', 'income', 'sales', 'revenue', 'receipt'}
    _DEBIT_SECTIONS  = {'asset', 'expenditure', 'expense', 'purchase'}
    _PL_SECTIONS     = {'profit', 'loss', 'trading', 'p & l', 'p&l', 'income', 'expenditure',
                        'sales', 'purchase', 'revenue', 'receipt', 'expense'}
    _INCOME_KW       = {'sales', 'income', 'revenue', 'receipt', 'interest received',
                        'commission received', 'discount received'}
    _EXPENSE_KW      = {'purchase', 'salary', 'rent', 'wages', 'depreciation', 'expense',
                        'expenditure', 'interest paid', 'commission paid', 'advertisement'}

    def _infer_pl_side(name: str) -> str:
        nl = name.lower()
        if any(kw in nl for kw in _INCOME_KW):
            return 'credit'
        if any(kw in nl for kw in _EXPENSE_KW):
            return 'debit'
        # Default P&L: left side (expenses) → debit
        return current_side

    for raw_line in text.split('\n'):
        line = raw_line.strip()
        if not line:
            continue

        ll = line.lower()

        # ── Detect section / group header ─────────────────────────────────────
        # A section header is a line with NO colon-amount pattern (or a known keyword line)
        m = _COLON_ENTRY_RE.match(line)
        if not m or not any(c.isdigit() for c in line):
            # Could be a section header
            matched_kw = next((kw for kw in list(_CREDIT_SECTIONS) + list(_DEBIT_SECTIONS) + list(_PL_SECTIONS)
                               if kw in ll), None)
            if matched_kw:
                in_pl = any(kw in ll for kw in _PL_SECTIONS - _CREDIT_SECTIONS - _DEBIT_SECTIONS)
                if any(kw in ll for kw in _CREDIT_SECTIONS):
                    current_side = 'credit'
                elif any(kw in ll for kw in _DEBIT_SECTIONS):
                    current_side = 'debit'
                # PROFIT & LOSS section — side will be inferred per-entry
                if any(kw in ll for kw in _PL_SECTIONS):
                    in_pl = True
                current_group = line.title()
            continue

        # ── Parse "Name : Amount" entry ───────────────────────────────────────
        name    = _clean(m.group(1).strip())
        amt_str = m.group(2).strip()
        if amt_str.startswith('(') and amt_str.endswith(')'):
            amt_str = '-' + amt_str[1:-1]
        amount = parse_amount(amt_str)

        if not name or amount == 0.0:
            continue
        if _is_skip(name):
            continue
        if len(name) < 2:
            continue

        side = _infer_pl_side(name) if in_pl else current_side

        e = TrialBalanceEntry(account_name=name, group=current_group)
        if side == 'credit':
            e.credit = abs(amount)
        else:
            e.debit  = abs(amount)
        entries.append(e)

    return entries


# ── Format I: Tally Single-Column Indented Trial Balance ──────────────────────

# Tally's standard (reserved) group names. A row bearing one of these names
# is structurally a GROUP even when the export indents it at the same level
# as its children (Tally does this for e.g. "Cash-in-hand" → "Cash").
_TALLY_STD_GROUPS = {
    'capital account', 'reserves & surplus', 'reserves and surplus',
    'loans (liability)', 'bank od a/c', 'bank occ a/c', 'secured loans',
    'unsecured loans', 'current liabilities', 'duties & taxes',
    'duties and taxes', 'provisions', 'sundry creditors',
    'fixed assets', 'investments', 'current assets', 'bank accounts',
    'cash-in-hand', 'cash in hand', 'deposits (asset)',
    'loans & advances (asset)', 'loans and advances (asset)',
    'stock-in-hand', 'stock in hand', 'sundry debtors',
    'branch / divisions', 'branch/divisions', 'misc. expenses (asset)',
    'misc expenses (asset)', 'suspense a/c', 'sales accounts',
    'purchase accounts', 'direct incomes', 'direct expenses',
    'indirect incomes', 'indirect expenses', 'opening stock',
    'closing stock', 'income (direct)', 'income (indirect)',
    'expenses (direct)', 'expenses (indirect)', 'retained earnings',
    # Common single-child group names found in Indian BS/P&L exports
    'purchase', 'purchases', 'purchase a/c', 'purchase ac',
    'sales', 'sale', 'sale a/c', 'sales a/c', 'sale ac',
    'direct expenditure', 'direct expinditure', 'direct expinditure',
    'indirect expenditure', 'indirect expinditure',
    'indirect expences', 'indirect expenses', 'direct expences',
    'expenses direct', 'expenses indirect',
    'current assets', 'current liabilities', 'fixed assets',
    'loans liabilities', 'loans & borrowings', 'loans and borrowings',
    'capital a/c', 'capital ac', 'sundry payables', 'sundry receivables',
    'broker master', 'broker a/c',
    'other income', 'other expenses', 'other expenditure',
}


def _is_tally_std_group(name: str) -> bool:
    return name.strip().lower() in _TALLY_STD_GROUPS


def _is_tally_std_group(name: str) -> bool:
    return name.strip().lower() in _TALLY_STD_GROUPS


def _fold_tally_hierarchy(
    raw: List[Tuple[int, str, float, float]],
    grand_dr: float = 0.0, grand_cr: float = 0.0,
    source: str = '',
    indents: Optional[List[Optional[float]]] = None,
) -> List[TrialBalanceEntry]:
    """
    Shared parent/leaf reconstruction for Tally hierarchical TB exports
    (XLS Format I and PDF Format I-PDF).

    raw: ordered (row_idx, name, dr, cr) rows, Grand Total row EXCLUDED.
    indents: optional per-row indent level (xlsx cell indent / PDF x-position).
        When available, a row may only absorb following rows whose indent is
        STRICTLY GREATER than its own — this disambiguates two adjacent equal
        sibling amounts (which would otherwise look like a parent+child pair)
        from a genuine single-child group.

    A Tally TB is a tree serialized depth-first (children follow their
    parent). Sub-groups carry their own Dr/Cr totals, so the reliable signal
    is arithmetic: a row is a parent iff the totals of the units that follow
    it sum to its own totals. We fold RIGHT-TO-LEFT so inner groups collapse
    into single units before their ancestors are tested — this handles
    arbitrary nesting depth and mixed Dr/Cr children (signed netting), which
    the old single-pass magnitude-sum approach could not.
    """
    n = len(raw)
    if n == 0:
        return []

    def _indent_of(i):
        if indents is None or i >= len(indents):
            return None
        return indents[i]

    # Each unit: {'idx', 'name', 'dr', 'cr', 'indent', 'children': [units]}
    units = [{'idx': i, 'name': raw[i][1], 'dr': raw[i][2], 'cr': raw[i][3],
              'indent': _indent_of(i), 'children': []} for i in range(n)]

    # Map row index -> position in `units` (positions shift as we fold, so we
    # re-locate by scanning; n is small enough that this stays fast).
    k = n - 2
    while k >= 0:
        # Find current position of original row k in units
        pos = next((p for p, u in enumerate(units) if u['idx'] == k), -1)
        if pos < 0 or pos == len(units) - 1:
            k -= 1
            continue
        u = units[pos]
        dr_i, cr_i = u['dr'], u['cr']
        if dr_i == 0.0 and cr_i == 0.0:
            k -= 1
            continue

        run_dr = run_cr = 0.0
        matched = -1
        ind_i = u['indent']
        for m in range(pos + 1, min(pos + 200, len(units))):
            run_dr += units[m]['dr']
            run_cr += units[m]['cr']

            # Indent veto (single-child folds only): two adjacent rows with
            # equal amounts at the same/lower indent are SIBLINGS, not a
            # parent+child pair. Tally indents are too noisy to constrain
            # multi-row sum matches (children sometimes share the parent's
            # indent), but a 1:1 equal-amount fold needs the child to be
            # strictly deeper.
            def _indent_ok():
                if m != pos + 1:
                    return True
                # Standard Tally group names are structural parents even at
                # the same indent (e.g. "Cash-in-hand" → "Cash").
                if _is_tally_std_group(u['name']):
                    return True
                ind_m = units[m]['indent']
                if ind_i is None or ind_m is None:
                    return True
                return ind_m > ind_i

            if dr_i != 0.0 and cr_i != 0.0:
                tol_d = max(1.0, dr_i * 0.002)
                tol_c = max(1.0, cr_i * 0.002)
                if (abs(run_dr - dr_i) <= tol_d and abs(run_cr - cr_i) <= tol_c
                        and _indent_ok()):
                    matched = m
                    break
                if run_dr > dr_i * 1.5 + 1000 and run_cr > cr_i * 1.5 + 1000:
                    break
            else:
                net_i = dr_i - cr_i
                tol = max(1.0, abs(net_i) * 0.002)
                if (abs((run_dr - run_cr) - net_i) <= tol
                        and (run_dr + run_cr) > 0 and _indent_ok()):
                    matched = m
                    break
                if (run_dr + run_cr) > (abs(net_i)) * 6 + 100000:
                    break
        if matched >= 0:
            u['children'] = units[pos + 1:matched + 1]
            del units[pos + 1:matched + 1]
        k -= 1

    # Flatten: emit leaves; group = innermost parent name.
    entries: List[TrialBalanceEntry] = []

    def _walk(unit, group):
        if unit['children']:
            for ch in unit['children']:
                _walk(ch, unit['name'])
            return
        dr, cr = unit['dr'], unit['cr']
        if dr == 0.0 and cr == 0.0:
            return                      # label-only row
        e = TrialBalanceEntry(account_name=unit['name'], group=group)
        if dr != 0.0 and cr != 0.0:
            # Unresolved parent (children didn't sum-match) — keep BOTH sides
            # so the file's totals are preserved; validator will flag it.
            e.debit, e.credit = dr, cr
        elif dr != 0.0:
            e.debit = dr
        else:
            e.credit = cr
        entries.append(e)

    for u in units:
        _walk(u, '')

    if grand_dr or grand_cr:
        sum_dr = round(sum(e.debit for e in entries), 2)
        sum_cr = round(sum(e.credit for e in entries), 2)
        if abs(sum_dr - grand_dr) > 1.0 or abs(sum_cr - grand_cr) > 1.0:
            logger.warning(
                "Tally TB fold mismatch in %s: leaves Dr=%.2f Cr=%.2f vs "
                "Grand Total Dr=%.2f Cr=%.2f", source, sum_dr, sum_cr,
                grand_dr, grand_cr)
        else:
            logger.info("Tally TB fold verified against Grand Total "
                        "(Dr=Cr=%.2f) for %s", grand_dr, source)
    return entries


# Layout (from TrialBal.xlsx sample):
#   Col0=Particulars, Col1=Debit closing, Col2=Credit closing
#   Parent rows have BOTH col1+col2 non-zero (group total in col2)
#   OR sum of following leaf rows equals the parent's single-col amount.
#   Leaf rows have ONLY col1 (Debit) or col2 (Credit) non-zero.

def _is_tally_single_col_tb(df: pd.DataFrame) -> bool:
    """Return True if df looks like a Tally 3-col indented Trial Balance."""
    for ri in range(min(12, len(df))):
        row = [str(c).strip().lower() for c in df.iloc[ri]]
        if not any('particular' in c for c in row):
            continue
        for ri2 in range(ri, min(ri + 5, len(df))):
            row2 = [str(c).strip().lower() for c in df.iloc[ri2]]
            if any('debit' in c or 'closing' in c for c in row2[1:3]):
                return True
    return False


def _parse_tally_single_col_tb(df: pd.DataFrame, source: str = '') -> List[TrialBalanceEntry]:
    """
    Parse Tally 3-column indented TB (Format I).

    col0 = Account / Group name
    col1 = Debit closing balance (leaf) or sub-total (parent)
    col2 = Credit closing balance (leaf) or group total (parent)

    Two-pass approach:
      Pass 1: Build flat list of (orig_row_idx, name, dr, cr)
      Pass 2: Mark rows as PARENT when:
        (a) Both col1 AND col2 are non-zero (signal A), OR
        (b) The row's amount equals the sum of immediately following
            single-col rows (signal B / lookahead sum match)
      Pass 3: Emit only LEAF rows as TB entries; parents set current_group.
    """
    # Locate the header row (has "particulars")
    hdr_row = -1
    for ri in range(min(15, len(df))):
        row = [str(c).strip().lower() for c in df.iloc[ri]]
        if any('particular' in c for c in row):
            hdr_row = ri
            break
    if hdr_row < 0:
        return []

    # ── Pass 1: collect raw rows (stop at Grand Total) ────────────────────────
    # Tally exports often repeat the final page block AFTER the Grand Total
    # row; everything past the first Grand Total is a duplicate and must be
    # ignored. The Grand Total itself is captured as the verification target.

    # Per-row indent levels (for fold disambiguation). For .xlsx we read the
    # cell's alignment.indent via openpyxl (df row i ↔ sheet row i+1, since
    # the df is loaded header=None). Fallback: leading spaces in the raw text.
    xlsx_indents: Optional[dict] = None
    if source and source.lower().endswith('.xlsx'):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(source, read_only=False, data_only=True)
            ws = wb.active
            xlsx_indents = {}
            for r_i, ws_row in enumerate(ws.iter_rows(min_col=1, max_col=1)):
                cell = ws_row[0]
                ind = cell.alignment.indent if cell.alignment else 0
                xlsx_indents[r_i] = float(ind or 0)
            wb.close()
        except Exception as exc:            # pragma: no cover
            logger.debug("indent extraction failed for %s: %s", source, exc)
            xlsx_indents = None

    raw: List[Tuple[int, str, float, float]] = []  # (orig_row_idx, name, dr, cr)
    indents: List[Optional[float]] = []
    grand_dr = grand_cr = 0.0
    for row_idx in range(hdr_row + 1, len(df)):
        row = list(df.iloc[row_idx])
        while len(row) < 3:
            row.append('')

        name_raw = str(row[0])
        if name_raw.strip().lower() in ('nan', ''):
            continue
        name = _clean(name_raw.strip())
        if not name:
            continue
        if re.match(r'^\d{1,2}[-/][A-Za-z]', name) or re.match(r'^\d{4}$', name):
            continue
        dr = parse_amount(str(row[1]).strip() if len(row) > 1 else '')
        cr = parse_amount(str(row[2]).strip() if len(row) > 2 else '')
        if 'grand total' in name.lower():
            grand_dr, grand_cr = dr, cr
            break                       # ignore duplicated trailing block
        if _is_total_row(name):
            continue
        if xlsx_indents is not None:
            ind = xlsx_indents.get(row_idx, None)
        else:
            lead = len(name_raw) - len(name_raw.lstrip(' '))
            ind = float(lead)
        raw.append((len(raw), name, dr, cr))
        indents.append(ind)

    if not raw:
        return []

    # Only pass indents if they actually vary — a flat all-equal indent list
    # would block ALL folding (every row would look like a sibling).
    distinct = {i for i in indents if i is not None}
    use_indents = indents if len(distinct) > 1 else None

    entries = _fold_tally_hierarchy(raw, grand_dr, grand_cr, source,
                                    indents=use_indents)
    logger.info("Format I (Tally single-col TB) parsed %d entries from %s",
                len(entries), source)
    return entries



# ── Format J: Two-sided section-aware PDF (Liability/Asset or P&L two-column) ─
#
# Handles PDFs like JAI KANHAIYA MILLS LLP where:
#   Left column  = Liability / Expenses (Credit for BS, Debit for P&L)
#   Right column = Asset / Income (Debit for BS, Credit for P&L)
#   Section header rows = lines with NO amount (e.g. "Capital Account",
#                         "Current Liabilities", "Sundry Creditors")
#   Entry rows = Name + Amount on same or adjacent rows
#
# Detection: page has "Liability"/"Asset" or "Account Name Amount" repeated
# on left AND right side header row.

_J_SKIP_KW = {
    'total', 'grand total', 'page', 'from', 'to', 'amount',
    'carried', 'over', 'brought', 'forward',
    'balance sheet', 'profit and loss', 'profit & loss',
    'end of report', 'reporting',
}
# Exact-match noise names (used with 'nl in _J_NOISE_NAMES' not substring)
_J_NOISE_NAMES = {
    'as per enclosed list', 'end of report', 'page 1 of 1', 'page 1 of',
    # NOTE: 'profit and loss a/c' and variants are NOT here — they may be real ledger
    # entries inside P&L statements (e.g. Net Profit row inside Indirect Expenses section).
    # They are handled by _normalize_profit_leaves after collapse.
    'gross profit b/d',
    'b/d', 'income', 'indirect', 'total :', 'total',
    'liability', 'liabilities', 'asset', 'assets',
    'account', 'account name', 'amount',
    'as per enclosed', 'enclosed list',
}
_J_SECTION_KW = {
    # BS Left (Credit / Liability)
    'capital a/c':                     'CAPITAL',
    'capital account':                 'CAPITAL',
    'partners capital':                'CAPITAL',
    'current liabilities':             'PROVISIONS',
    'current liabities':               'PROVISIONS',
    'duties and taxes':                'DUTIES AND TAXES',
    'brokers (sale commission agents)':'SUNDRY CREDITORS',
    'provision':                       'PROVISIONS',
    'sundry creditors':                'SUNDRY CREDITORS',
    'sundry creditors mill':           'SUNDRY CREDITORS',
    'sundry payables':                 'SUNDRY CREDITORS',
    'loans liabilities':               'UNSECURED LOANS',
    'loans & borrowings':              'UNSECURED LOANS',
    'unsecured loans':                 'UNSECURED LOANS',
    # BS Right (Debit / Asset)
    'current assets':                  'OTHER CURRENT ASSETS',
    'current assests':                 'OTHER CURRENT ASSETS',
    'cash & bank balances':            'CASH AND BANK',
    'cash and bank balances':          'CASH AND BANK',
    'bank a/c':                        'CASH AND BANK',
    'loans / advances a/c':            'LOANS AND ADVANCES (ASSETS)',
    'loans and advances':              'LOANS AND ADVANCES (ASSETS)',
    'loans and advances a/c':          'LOANS AND ADVANCES (ASSETS)',
    'sundry debtors':                  'SUNDRY DEBTORS',
    'sundry receivables':              'SUNDRY DEBTORS',
    'fixed assets':                    'FIXED ASSETS',
    'deposits and investments':        'INVESTMENTS',
    'other current assets':            'OTHER CURRENT ASSETS',
    'stock in hand':                   'CLOSING STOCK',
    'closing stock':                   'CLOSING STOCK',
    'investments':                     'INVESTMENTS',
    'deposits':                        'DEPOSITS',
    'cash in hand':                    'CASH IN HAND',
    # P&L Left (Debit / Expenses)
    'opening stock':                   'OPENING STOCK',
    'purchase':                        'PURCHASE A/C',
    'direct expenses':                 'DIRECT EXPENSES (M)',
    'indirect expenses':               'INDIRECT EXPENSES',
    'indirect expences':               'INDIRECT EXPENSES',
    'expenses direct':                 'DIRECT EXPENSES (M)',
    'expenses direct (t&m)':           'DIRECT EXPENSES (M)',
    'expenses indirect':               'INDIRECT EXPENSES',
    'expenses indirect (p&l)':         'INDIRECT EXPENSES',
    # P&L Right (Credit / Income)
    'sales':                           'SALES A/C',
    'gross profit':                    'CAPITAL',
    'indirect income':                 'INDIRECT INCOMES',
    'indirect incomes':                'INDIRECT INCOMES',
    'other income':                    'INDIRECT INCOMES',
}


def _detect_two_sided_section_pdf(file_path: str):
    """
    Return (split_x, is_pl, left_amt_x, right_amt_x) if this PDF is a
    two-sided section-header BS or P&L, else None.

    Detects by finding:
      • "Liability" + "Asset" headers (BS), OR
      • "Account Name Amount" twice on page 1 header (P&L)
    """
    try:
        import pdfplumber
        with pdfplumber.open(file_path) as pdf:
            if not pdf.pages:
                return None
            p0 = pdf.pages[0]
            words = p0.extract_words()
            text_lower = (p0.extract_text() or '').lower()

            # Check for Liability/Asset BS
            has_liab = 'liability' in text_lower or 'liabilities' in text_lower
            has_asset = 'asset' in text_lower or 'assets' in text_lower

            # Determine document type from the TITLE area (first 8 lines) only.
            # Body text may contain "Profit & Loss A/c" as an account name even
            # in a Balance Sheet — checking full page text causes false positives.
            page_lines = [l.strip() for l in (p0.extract_text() or '').split('\n') if l.strip()]
            title_text = ' '.join(page_lines[:8]).lower()

            # Explicit Balance Sheet title → never is_pl
            is_bs_title = ('balance sheet' in title_text
                           or 'balance-sheet' in title_text)
            # P&L detected only in title area
            is_pl = (not is_bs_title and (
                'profit and loss' in title_text
                or 'profit & loss' in title_text
                or 'trading' in title_text
            ))

            if not ((has_liab and has_asset) or is_pl):
                return None

            # Must NOT have a "Debit"+"Credit" header pair — those are Format H/I (TB columns).
            # "Particulars" alone is fine (two-sided P&L statements use it as a column label).
            if 'debit' in text_lower and 'credit' in text_lower:
                return None

            page_w = float(p0.width)

            # ── Amount-gap approach to split_x ────────────────────────────────
            # A two-sided PDF has two distinct amount columns: one in the left
            # half and one in the right half.  We sample all pages and find the
            # maximum x of left-half amounts and the minimum x of right-half
            # amounts; split_x = midpoint of those two — far more robust than
            # using the x of section-header words (which sit inside the text
            # region, not at the true column boundary).
            _AMT_RE = re.compile(r'^-?[\d,]+\.?\d*$')

            all_amt_xs: List[float] = []
            for page in pdf.pages[:5]:          # sample up to 5 pages
                for w in page.extract_words():
                    txt = w['text'].replace(',','').lstrip('-').lstrip('(').rstrip(')')
                    if _AMT_RE.match(txt):
                        all_amt_xs.append(float(w['x0']))

            if not all_amt_xs:
                return None

            # Partition by page midpoint
            mid = page_w * 0.50
            left_xs  = [x for x in all_amt_xs if x < mid]
            right_xs = [x for x in all_amt_xs if x > mid]

            if not left_xs or not right_xs:
                # Fall back to 40/60 split of page
                split_x   = page_w * 0.50
                left_amt_x  = page_w * 0.40
                right_amt_x = page_w * 0.85
                return split_x, is_pl, left_amt_x, right_amt_x

            left_max  = max(left_xs)
            right_min_amt = min(right_xs)

            # Find the x of the leftmost RIGHT-SIDE TEXT word (section header or
            # account name on the right column). This is the true start of the
            # right half and gives a tighter split boundary than the midpoint of
            # left/right amounts (which places split_x too far into the page,
            # absorbing right-side text into the left name cluster).
            all_text_xs: List[float] = []
            for page in pdf.pages[:5]:
                for w in page.extract_words():
                    txt = w['text'].replace(',','').lstrip('-').lstrip('(').rstrip(')')
                    if not _AMT_RE.match(txt) and len(w['text']) > 1:
                        all_text_xs.append(float(w['x0']))

            # Right-side text = text words with x > left_max + 20pt gap
            right_text_xs = [x for x in all_text_xs if x > left_max + 20]
            right_text_min = min(right_text_xs) if right_text_xs else right_min_amt

            # Place split_x in the gap between end of left amounts and start of right text
            # Use the smaller of: midpoint(left_max, right_text_min) or midpoint(left_max, right_min_amt)
            split_x = min(
                (left_max + right_text_min) / 2.0,
                (left_max + right_min_amt) / 2.0
            )

            left_amt_x  = left_max
            right_amt_x = max(right_xs)

            return split_x, is_pl, left_amt_x, right_amt_x
    except Exception:
        return None


def _parse_two_sided_section_pdf(file_path: str) -> List[TrialBalanceEntry]:
    """
    Parse two-sided section-aware BS or P&L PDF (Format J).

    Algorithm per page:
      1. Cluster words into rows by y-position (3pt tolerance).
      2. Split each row into LEFT side (x < split_x) and RIGHT side (x >= split_x).
      3. For each side independently:
         a. If a row has a name but NO amount → it's a section header;
            set current_group for that side.
         b. If a row has name + amount → emit a leaf entry under current_group.
         c. If a row has ONLY an amount (name on previous row) → attach to prev name.
    """
    import pdfplumber

    layout = _detect_two_sided_section_pdf(file_path)
    if not layout:
        return []

    split_x, is_pl, left_amt_x, right_amt_x = layout

    # For BS: Left=Liability(Cr), Right=Asset(Dr)
    # For P&L: Left=Expenses(Dr), Right=Income(Cr)
    left_is_dr  = is_pl          # P&L left=expenses=Dr; BS left=liability=Cr
    right_is_dr = not is_pl      # P&L right=income=Cr;  BS right=asset=Dr

    entries: List[TrialBalanceEntry] = []
    left_group = right_group = ''
    left_pending_name = right_pending_name = ''
    left_orphan_amt: Optional[float] = None   # amount seen before its name
    right_orphan_amt: Optional[float] = None

    _NUM_PAT = re.compile(r'^-?[\d,]+\.?\d*$')

    def _is_num(s: str) -> bool:
        t = s.strip()
        # Handle Indian "(-)" negative prefix: "(-)3,59,608.65"
        if t.startswith('(-)'):
            t = t[3:]
        # Only strip a leading '(' if there is a matching trailing ')'
        elif t.startswith('(') and t.endswith(')'):
            t = t[1:-1]
        elif t.endswith(')') and not t.startswith('('):
            return False   # trailing ')' without opener → part of a name like "Mst 40)"
        t = t.lstrip('-').replace(',', '')
        return bool(_NUM_PAT.match(t))

    def _parse_num(s: str) -> float:
        t = s.strip()
        neg = t.startswith('-') or t.startswith('(-)')
        if t.startswith('(-)'):
            t = t[3:]
        elif t.startswith('(') and t.endswith(')'):
            t = t[1:-1]
        cleaned = t.lstrip('-').replace(',', '')
        try:
            val = float(cleaned)
            return -val if neg else val
        except ValueError:
            return 0.0

    def _emit(name: str, group: str, amt: float, is_dr: bool) -> None:
        if not name or abs(amt) < 0.01:
            return
        nl = name.lower().strip()
        if nl in _J_NOISE_NAMES:
            return
        if _is_total_row(name):
            return
        if re.match(r'^page\s+\d', nl):
            return
        e = TrialBalanceEntry(account_name=name, group=group)
        if is_dr:
            e.debit  = abs(amt) if amt >= 0 else 0.0
            e.credit = abs(amt) if amt <  0 else 0.0
        else:
            e.credit = abs(amt) if amt >= 0 else 0.0
            e.debit  = abs(amt) if amt <  0 else 0.0
        entries.append(e)

    def _resolve_section(raw_name: str, side: str) -> str:
        """Map a section header text to a canonical group name."""
        nl = raw_name.lower().strip()
        # Exact match first
        if nl in _J_SECTION_KW:
            return _J_SECTION_KW[nl]
        # Prefix match
        for kw, grp in _J_SECTION_KW.items():
            if nl.startswith(kw) or kw.startswith(nl[:10]):
                return grp
        # Fallback: use raw name uppercased
        return raw_name.strip().upper()

    # First data row threshold — skip company header rows (first ~60pt of page)
    # These contain company name, title, date etc., not financial data.
    # The actual column headers (LIABILITIES/ASSETS or Account Name/Amount) appear
    # at y≈54, so data rows start at y≈70+.
    HEADER_AREA_TOP = 60.0  # rows at or above this y are pure header noise

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(keep_blank_chars=False)
            if not words:
                continue

            # Cluster into rows with 2pt tolerance, then merge any adjacent
            # name-only and amount-only clusters within 4pt — this handles
            # PDFs where some rows have the amount word rendered 0.2–1.5pt
            # above or below the text words (causing them to fall in different
            # 2pt buckets while actually belonging to the same logical row).
            words.sort(key=lambda w: (round(float(w['top']) / 2) * 2, float(w['x0'])))
            rows: Dict[int, list] = {}
            for w in words:
                y = round(float(w['top']) / 2) * 2
                rows.setdefault(y, []).append(w)

            # Post-merge: absorb isolated name-only or amount-only rows into
            # their nearest neighbor within 4pt, provided merging would supply
            # what the neighbor lacks.
            ys = sorted(rows.keys())
            merged: Dict[int, list] = {}
            skip_ys: set = set()
            for idx, y in enumerate(ys):
                if y in skip_ys:
                    continue
                row_words = rows[y]
                has_name = any(not _is_num(w['text']) for w in row_words)
                has_amt  = any(_is_num(w['text']) for w in row_words)
                # Look at next bucket
                if idx + 1 < len(ys):
                    y2 = ys[idx + 1]
                    if y2 - y <= 4:
                        row2 = rows[y2]
                        has_name2 = any(not _is_num(w['text']) for w in row2)
                        has_amt2  = any(_is_num(w['text']) for w in row2)
                        # Merge if one has only names and the other has only amounts
                        if (has_name and not has_amt and has_amt2 and not has_name2) or \
                           (has_amt and not has_name and has_name2 and not has_amt2):
                            merged[y] = row_words + row2
                            skip_ys.add(y2)
                            continue
                merged[y] = row_words
            rows = merged

            for y in sorted(rows.keys()):
                # Skip page header area (company name, title, date rows)
                if y < HEADER_AREA_TOP:
                    continue

                rw = sorted(rows[y], key=lambda w: float(w['x0']))
                left_w  = [w for w in rw if float(w['x0']) < split_x]
                right_w = [w for w in rw if float(w['x0']) >= split_x]

                for side_w, side_name, is_dr, grp_ref, pend_ref in [
                    (left_w,  'L', left_is_dr,  'left_group',  'left_pending_name'),
                    (right_w, 'R', right_is_dr, 'right_group', 'right_pending_name'),
                ]:
                    if not side_w:
                        continue

                    # Separate name tokens from amount tokens in this side's words
                    name_tokens = [w['text'] for w in side_w if not _is_num(w['text'])]
                    amt_tokens  = [w for w in side_w if _is_num(w['text'])]

                    name_str = _clean(' '.join(name_tokens)).strip()
                    nl = name_str.lower()

                    # Skip header/footer/noise rows.
                    # "profit and loss" / "profit & loss" / "trading profit" skip only
                    # when there is NO amount on this row — those prefixes are used for
                    # document titles and column headers, but also appear as account names
                    # (e.g. "Profit And Loss A/c 763089.29") that must be kept.
                    _unconditional_skip = (
                        'account name', 'account amount',
                        'amount', 'balance sheet',
                        'grand total', 'end of report',
                        'reporting date', 'page 1 of', 'page no',
                        'liabilities amount', 'assets amount',
                    )
                    _pl_title_skip = ('profit and loss', 'profit & loss', 'trading profit')
                    if any(nl.startswith(sk) for sk in _unconditional_skip):
                        continue
                    if not amt_tokens and any(nl.startswith(sk) for sk in _pl_title_skip):
                        continue  # title/column-header row only; keep if it has an amount
                    # Skip standalone header words (exact match only)
                    if nl in ('liabilities', 'liability', 'assets', 'asset',
                              'account', 'amount', 'account name amount'):
                        continue
                    # Skip phone / fax number rows (all-digit name with dashes, or bare digits)
                    if name_str and re.match(r'^\d[\d\-]+$', name_str):
                        continue

                    cur_grp = left_group if side_name == 'L' else right_group

                    if not name_str:
                        # Amount-only row — no name text at all on this side
                        if amt_tokens:
                            # Attach to pending name if there is one
                            pend = left_pending_name if side_name == 'L' else right_pending_name
                            if pend:
                                if side_name == 'L':
                                    best = min(amt_tokens, key=lambda w: abs(float(w['x0']) - left_amt_x))
                                else:
                                    best = min(amt_tokens, key=lambda w: abs(float(w['x0']) - right_amt_x))
                                amt = _parse_num(best['text'])
                                _emit(pend, cur_grp, amt, is_dr)
                                if side_name == 'L':
                                    left_pending_name = ''
                                else:
                                    right_pending_name = ''
                            else:
                                # No pending name — save as orphan in case the name comes next row
                                if side_name == 'L':
                                    best = min(amt_tokens, key=lambda w: abs(float(w['x0']) - left_amt_x))
                                    left_orphan_amt = _parse_num(best['text'])
                                else:
                                    best = min(amt_tokens, key=lambda w: abs(float(w['x0']) - right_amt_x))
                                    right_orphan_amt = _parse_num(best['text'])
                        continue

                    # We have a name — check if it's a section header or a leaf entry
                    # If the name is a known section keyword, always update the group.
                    is_section_name = (nl in _J_SECTION_KW or
                                       any(nl.startswith(kw) for kw in _J_SECTION_KW))

                    if not amt_tokens:
                        # No amount → section header or name waiting for orphan amount
                        # Check if a preceding amount-only row on this side is waiting
                        orphan = left_orphan_amt if side_name == 'L' else right_orphan_amt
                        _is_noise_hdr = (
                            _is_total_row(name_str)
                            or _is_profit_row(name_str)
                            or nl in _J_NOISE_NAMES
                            or 'as per enclosed' in nl
                            or 'enclosed list' in nl
                            or re.match(r'^\d{1,2}/\d{2}/\d{4}', name_str)
                            or any(kw == nl for kw in _J_SKIP_KW)
                        )
                        if not _is_noise_hdr:
                            if orphan is not None:
                                # Preceding row had an amount with no name — attach it here
                                _emit(name_str, cur_grp, orphan, is_dr)
                                if side_name == 'L':
                                    left_orphan_amt = None
                                else:
                                    right_orphan_amt = None
                            else:
                                canonical = _resolve_section(name_str, side_name)
                                if side_name == 'L':
                                    left_group = canonical
                                    left_pending_name = name_str
                                else:
                                    right_group = canonical
                                    right_pending_name = name_str
                        elif orphan is not None:
                            # Noise name — clear the orphan rather than leaving it dangling
                            if side_name == 'L':
                                left_orphan_amt = None
                            else:
                                right_orphan_amt = None
                    else:
                        # Has amounts
                        if side_name == 'L':
                            best = min(amt_tokens, key=lambda w: abs(float(w['x0']) - left_amt_x))
                        else:
                            best = min(amt_tokens, key=lambda w: abs(float(w['x0']) - right_amt_x))
                        amt = _parse_num(best['text'])

                        _is_noise_entry = (
                            _is_total_row(name_str)
                            or _GROSS_PROFIT_RE.search(nl) is not None
                            or re.match(r'^page\s+\d', nl)
                            or nl in _J_NOISE_NAMES
                            or nl in ('b/d', 'income', 'indirect', 'claim', 'a/c',
                                      'and', 'or', '&', 'on sale', 'on purchase',
                                      'total', 'total :', 'grand total')
                            or 'as per enclosed' in nl
                            or 'enclosed list' in nl
                            or re.match(r'^\d{1,2}/\d{2}/\d{4}', name_str)
                            or any(kw == nl for kw in _J_SKIP_KW)
                        )

                        if name_str and not _is_noise_entry:
                            # If section keyword with amount: update group then emit
                            if is_section_name:
                                canonical = _resolve_section(name_str, side_name)
                                if side_name == 'L':
                                    left_group = canonical
                                else:
                                    right_group = canonical
                            cur_grp = left_group if side_name == 'L' else right_group
                            _emit(name_str, cur_grp, amt, is_dr)
                            if side_name == 'L':
                                left_pending_name = ''
                                left_orphan_amt = None   # clear any stale orphan
                            else:
                                right_pending_name = ''
                                right_orphan_amt = None

    # ── Post-pass: collapse subtotal/group rows that also carried amounts ────
    # Section headers with printed totals (and schedule subtotals) get emitted
    # alongside their detail rows, double counting. We collapse per-side using
    # SIGNED amounts (negative entries stay in the same side they were emitted on
    # rather than being flipped to the other side). This lets the collapse engine
    # see Purchase(+2154M) alongside its negative-amount children like
    # Yarn Purchase Return(-4.57M) and sum-match correctly.
    final: List[TrialBalanceEntry] = []
    for side_is_dr in (True, False):
        # Collect entries that PRIMARILY belong to this side (debit entries for Dr pass,
        # credit entries for Cr pass) — but also include same-side negative entries
        # (entries whose "primary" side is this side but whose amount ended up negative,
        # flipping them to the other side in _emit). We detect these by checking if they
        # are listed with debit=0,credit>0 but the PDF row was on the Dr/left side; we
        # can't recover the original side after _emit without tracking it, so instead we
        # use a simpler invariant: on each side pass, collect ALL remaining entries
        # as signed amounts, collapse, then claim the output for this side.
        #
        # More precisely: on the Dr pass, collect everything that has debit>0 (positive Dr)
        # plus credit entries that were emitted as Cr due to sign-flip of a Dr-side amount.
        # Since we can't distinguish these after _emit, we handle them via the signed approach:
        # for the Dr pass, include entries where debit > 0; for the Cr pass, entries where
        # credit > 0. Negative children (flipped to Cr) will appear in the Cr pass at their
        # correct magnitude. The collapse for the Dr pass may still not sum-match if negatives
        # are missing, but this is an inherent limitation of per-side processing.
        # The Kaliya and JHANWAR PDFs that formerly worked still work here.
        side_rows = [e for e in entries
                     if (side_is_dr and e.debit > 0) or (not side_is_dr and e.credit > 0)]
        pairs = [(e.account_name, e.debit if side_is_dr else e.credit)
                 for e in side_rows]
        groups = {e.account_name: e.group for e in side_rows}
        leaves = _normalize_profit_leaves(_collapse_side_pairs(pairs))
        for name, amt, grp in leaves:
            e = TrialBalanceEntry(account_name=name,
                                  group=grp or groups.get(name, ''))
            if side_is_dr:
                e.debit = abs(amt)
            else:
                e.credit = abs(amt)
            final.append(e)
    entries = final

    logger.info("Format J (two-sided section PDF) parsed %d entries from %s",
                len(entries), file_path)
    return entries


# ── Format I-PDF: Tally Single-Column Trial Balance PDF ───────────────────────
# Identical structure to the XLS Format I but extracted from PDF word positions.
# Page layout:
#   Name words at x < dr_col_x
#   Debit amount word right-aligned near x ≈ dr_col_x
#   Credit amount word right-aligned near x ≈ cr_col_x
# Detection: page 1 has both "Debit" and "Credit" headers on the same line,
# with NO "PARTICULARS" / "AMOUNT" two-column header (which Format H covers).

_IND_NUM_RE = re.compile(
    # Matches: optional negative/paren, digits with Indian comma grouping, optional decimal
    # Handles: 7,06,846.00  90,00,404.29  2,52,65,556.00  1,37,21,110.00  37,93,384.76
    r'^\(?-?\d{1,3}(?:,\d{2,3})*(?:\.\d+)?\)?$'
)


def _parse_indian_amount(text: str) -> float:
    """Parse Indian lakh-formatted number like 90,00,404.29 → 9000404.29."""
    t = text.strip().lstrip('(').rstrip(')')
    negative = text.strip().startswith('(') and text.strip().endswith(')')
    # Remove ALL commas — Indian format uses commas at 2-digit groups
    t = t.replace(',', '')
    try:
        val = float(t)
        return -val if negative else val
    except ValueError:
        return 0.0


def _detect_tally_single_col_tb_pdf(file_path: str):
    """Return (debit_x, credit_x) split point or None if not this format."""
    try:
        import pdfplumber
        with pdfplumber.open(file_path) as pdf:
            if not pdf.pages:
                return None
            words = pdf.pages[0].extract_words()
            # Find "Debit" and "Credit" header words
            dr_word = next((w for w in words if w['text'].strip().lower() == 'debit'), None)
            cr_word = next((w for w in words if w['text'].strip().lower() == 'credit'), None)
            if not dr_word or not cr_word:
                return None
            dr_x = float(dr_word['x0'])
            cr_x = float(cr_word['x0'])
            # Sanity: credit header must be to the right of debit header
            if cr_x <= dr_x:
                return None
            # The split between Dr and Cr amounts is midway between headers
            split_x = (dr_x + cr_x) / 2.0
            return dr_x, cr_x, split_x
    except Exception:
        return None


def _parse_tally_single_col_pdf(file_path: str) -> List[TrialBalanceEntry]:
    """
    Parse Tally-exported single-column Trial Balance PDF (Format I-PDF).

    Uses word x-positions to assign each amount to Debit or Credit column,
    then applies the same two-pass parent/leaf detection as the XLS Format I.

    Indian lakh format (90,00,404.29) is handled by _parse_indian_amount.
    """
    import pdfplumber

    layout = _detect_tally_single_col_tb_pdf(file_path)
    if not layout:
        return []

    dr_header_x, cr_header_x, split_x = layout

    # ── Pass 1: collect (name, dr, cr) for every row across all pages ─────────
    raw: List[Tuple[int, str, float, float]] = []
    indents: List[Optional[float]] = []
    grand_dr = grand_cr = 0.0
    hit_grand = False
    order = 0

    _SKIP_TEXT = {
        'debit', 'credit', 'closing', 'balance', 'particulars',
        'page', 'carried', 'over', 'continued', 'brought', 'forward',
        'trial',
    }

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            if hit_grand:
                break
            words = page.extract_words(keep_blank_chars=False)
            if not words:
                continue

            # Cluster words into rows by y-position (tolerance 3pt)
            words.sort(key=lambda w: (round(float(w['top']) / 3) * 3, float(w['x0'])))
            rows: Dict[int, list] = {}
            for w in words:
                y_key = round(float(w['top']) / 3) * 3
                rows.setdefault(y_key, []).append(w)

            for y_key in sorted(rows.keys()):
                row_w = sorted(rows[y_key], key=lambda w: float(w['x0']))

                # Name words are left of the Debit column (with 60pt margin)
                # Amount words match the Indian number pattern
                name_cutoff = dr_header_x - 60  # well left of first amount column
                name_words = [w for w in row_w if float(w['x0']) < name_cutoff
                              and not _IND_NUM_RE.match(w['text'].replace(' ',''))]
                amt_words   = [w for w in row_w
                               if _IND_NUM_RE.match(w['text'].replace(' ',''))]

                if not amt_words and not name_words:
                    continue

                # Build name from left-side text words
                name = _clean(' '.join(w['text'] for w in name_words).strip())
                if not name:
                    continue
                nl = name.lower().strip()
                # Grand Total = verification target; everything after it is a
                # duplicated trailing block — stop here.
                if 'grand total' in nl:
                    for aw in amt_words:
                        amt = _parse_indian_amount(aw['text'])
                        if float(aw['x0']) < split_x:
                            grand_dr = amt
                        else:
                            grand_cr = amt
                    hit_grand = True
                    break
                # Skip header/footer rows (whole-word match — substring
                # matching would drop e.g. 'Packing Matrial Gst' via 'trial')
                nl_words = set(re.findall(r'[a-z]+', nl))
                if nl_words & _SKIP_TEXT:
                    continue
                if _is_total_row(name):
                    continue
                if re.match(r'^\d{1,2}[-/][A-Za-z]', name):
                    continue

                # Assign each amount word to Dr or Cr by x-position
                dr_val = cr_val = 0.0
                for aw in amt_words:
                    ax = float(aw['x0'])
                    amt = _parse_indian_amount(aw['text'])
                    if ax < split_x:
                        dr_val = amt   # left of split → Debit
                    else:
                        cr_val = amt   # right of split → Credit

                raw.append((order, name, dr_val, cr_val))
                # Indent proxy: x-position where the name starts
                indents.append(float(name_words[0]['x0']) if name_words else None)
                order += 1

    if not raw:
        return []

    # Quantize x-positions into indent levels (3pt tolerance) so equal-level
    # siblings compare as exactly equal in the fold engine.
    q_indents: List[Optional[float]] = [
        None if x is None else round(x / 3.0) for x in indents]
    distinct = {i for i in q_indents if i is not None}
    use_indents = q_indents if len(distinct) > 1 else None

    # Re-index sequentially (page clustering may have gaps) and fold.
    raw = [(i, name, dr, cr) for i, (_, name, dr, cr) in enumerate(raw)]
    entries = _fold_tally_hierarchy(raw, grand_dr, grand_cr, source=file_path,
                                    indents=use_indents)
    logger.info("Format I-PDF parsed %d entries from %s", len(entries), file_path)
    return entries


# ── Main parser ───────────────────────────────────────────────────────────────

class TrialBalanceParser:

    def parse_file(self, file_path: str) -> List[TrialBalanceEntry]:
        # Normalise extension — handles .XLS, .XLSX uppercase, no-extension, etc.
        raw_ext = file_path.rsplit('.', 1)[-1].lower() if '.' in file_path else ''
        # Strip query-string artefacts like "file.xlsx?t=123"
        raw_ext = raw_ext.split('?')[0].split('&')[0].strip()

        if raw_ext == 'pdf':
            return self._parse_pdf(file_path)

        df = self._read_tabular(file_path, raw_ext)
        df = df.fillna('').map(lambda x: str(x).strip())
        return self._parse_dataframe(df, file_path)

    def _read_tabular(self, file_path: str, ext: str) -> pd.DataFrame:
        """Read any tabular file, trying multiple engines so we never reject a
        format outright. For unknown/missing extensions, ALL strategies are tried."""
        if ext in ('xlsx', 'xlsm', 'ods'):
            attempts = [('excel', 'openpyxl'), ('excel', 'xlrd'), ('html', None), ('csv', None)]
        elif ext == 'xls':
            attempts = [('excel', 'xlrd'), ('excel', 'openpyxl'), ('html', None), ('csv', None)]
        elif ext in ('csv', 'tsv', 'txt'):
            attempts = [('csv', None), ('excel', 'xlrd'), ('html', None)]
        else:
            # Unknown or missing extension — try every strategy
            attempts = [('excel', 'openpyxl'), ('excel', 'xlrd'), ('html', None), ('csv', None)]

        last_err = None
        for kind, engine in attempts:
            try:
                if kind == 'excel':
                    return pd.read_excel(file_path, header=None, dtype=str, engine=engine)
                if kind == 'html':
                    tables = pd.read_html(file_path)  # returns list of DataFrames
                    if tables:
                        return max(tables, key=lambda t: t.shape[0]).astype(str)
                if kind == 'csv':
                    sep = '\t' if ext == 'tsv' else None
                    return pd.read_csv(file_path, header=None, dtype=str, sep=sep,
                                       engine='python', on_bad_lines='skip')
            except Exception as e:  # try the next strategy
                last_err = e
                continue
        logger.warning("All read strategies failed for '%s': %s — returning empty DataFrame",
                       file_path, last_err)
        return pd.DataFrame()

    def _parse_dataframe(self, df: pd.DataFrame, source: str = '') -> List[TrialBalanceEntry]:
        # Strategy I: Tally single-column indented TB (3-col: Name | Debit | Credit)
        # Check FIRST because these files have "Particulars" header but NO separate
        # Dr/Cr column header for the credit side (only "Debit" is labeled), which
        # causes Format A to miss them.
        if _is_tally_single_col_tb(df):
            entries = _parse_tally_single_col_tb(df, source)
            if entries:
                logger.info("Parsed %d TB entries (Format I / Tally single-col) from %s",
                            len(entries), source)
                return entries

        # Strategy A: Standard Dr/Cr header
        header_row, col_map = _detect_columns(df)
        if header_row >= 0:
            entries = self._extract_standard(df, header_row, col_map)
            if entries:
                logger.info("Parsed %d TB entries (standard) from %s", len(entries), source)
                return entries

        # Strategy C (preferred): two-sided PARTICULARS/Account + Amount layout.
        # This handles both indented and flat group/leaf structures and feeds
        # the master grouping dictionary, so prefer it when a clear two-sided
        # header is present.
        if _is_particulars_two_column(df):
            entries = _parse_particulars_two_column(df, is_pl_hint=_filename_is_pl(source))
            if entries:
                logger.info("Parsed %d TB entries (two-column) from %s", len(entries), source)
                return entries

        # Strategy B: Balance Sheet Liabilities/Assets layout
        if _is_balance_sheet_format(df):
            entries = _parse_balance_sheet_format(df)
            if entries:
                logger.info("Parsed %d TB entries (BS/PL) from %s", len(entries), source)
                return entries

        logger.warning(
            "Could not detect column layout in '%s'. "
            "Tried: Format I (Tally single-col TB), Standard Dr/Cr, "
            "Two-column PARTICULARS/AMOUNT, Balance Sheet Liabilities/Assets. "
            "Returning empty — file will be skipped with a warning.", source
        )
        return []

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

            # Totals are always derived rows — skip.
            # Profit/Loss rows are skipped ONLY when they carry no amount:
            # a standard TB can legitimately contain a "NET PROFIT" ledger
            # (grouped under CAPITAL, Dr) which must be preserved.
            if _is_total_row(account_val):
                continue
            if _is_profit_row(account_val) and debit == 0.0 and credit == 0.0:
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

        # ── Try Format J: Two-sided section-aware BS/P&L PDF ───────────────────
        # Handles PDFs with "Liability/Asset" or "Account Name Amount" two-column
        # layout with section headers (Capital Account, Current Liabilities, etc.)
        # MUST run before Format H (which also matches "particulars" PDFs) and
        # before Format I-PDF (which matches "Debit/Credit" column header PDFs).
        try:
            j_entries = _parse_two_sided_section_pdf(file_path)
            if j_entries:
                logger.info("Parsed %d entries via two-sided section PDF (Format J)",
                            len(j_entries))
                return j_entries
        except Exception as _je:
            logger.warning("Format J PDF parse failed: %s", _je)

        # ── Try Format I (PDF): Tally single-column Trial Balance ────────────────
        # Detects by finding "Debit" and "Credit" column headers on the same row.
        # Same logic as the XLS Format I but using word x-positions to split Dr/Cr.
        # Must run FIRST because this format has "Particulars" in the text which
        # would otherwise trigger Format H or table-extraction paths incorrectly.
        try:
            i_entries = _parse_tally_single_col_pdf(file_path)
            if i_entries:
                logger.info("Parsed %d entries via Tally single-col TB PDF (Format I-PDF)",
                            len(i_entries))
                return i_entries
        except Exception as _ie:
            logger.warning("Format I-PDF parse failed: %s", _ie)

        # ── Try Format H: PARTICULARS/AMOUNT two-sided (indentation-based) ──────
        # Runs FIRST: when a PARTICULARS/AMOUNT header exists, the indentation
        # method is the only one that correctly separates group headers from
        # detail rows (they share one amount column, so Format E double-counts
        # and table extraction merges sides).
        try:
            if _detect_particulars_two_col(file_path):
                h_entries = _parse_particulars_two_col_pdf(file_path)
                if h_entries:
                    logger.info("Parsed %d entries via PARTICULARS two-col "
                                "indentation format (Format H)", len(h_entries))
                    return h_entries
        except Exception as _he:
            logger.warning("Format H PDF parse failed: %s", _he)

        # ── Try Format G: Section-header "Name : Amount" BS/PL extract ──────────
        try:
            import pdfplumber as _ppl
            with _ppl.open(file_path) as _pdf:
                _text = '\n'.join(p.extract_text() or '' for p in _pdf.pages)
            if _is_section_header_format(_text):
                g_entries = _parse_section_header_format(_text)
                if g_entries:
                    logger.info("Parsed %d entries via section-header format (Format G)", len(g_entries))
                    return g_entries
        except Exception as _ge:
            logger.warning("Section-header PDF parse failed: %s", _ge)

        # ── Try Format F: Flat single-column "Name - Debit/Credit Amount" ────────
        try:
            import pdfplumber as _ppl
            with _ppl.open(file_path) as _pdf:
                _text = '\n'.join(p.extract_text() or '' for p in _pdf.pages)
            if _is_flat_line_format(_text):
                f_entries = _parse_flat_line_text(_text)
                if f_entries:
                    logger.info("Parsed %d entries via flat-line format (Format F)", len(f_entries))
                    return f_entries
        except Exception as _fe:
            logger.warning("Flat-line PDF parse failed: %s", _fe)

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
