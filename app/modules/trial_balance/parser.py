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


def _is_skip(name: str) -> bool:
    """Return True only for pure total rows. Profit/Loss rows are NOT skipped
    here — callers must check _is_profit_row separately and skip only when amount=0."""
    nl = name.lower().strip()
    return any(kw in nl for kw in _SKIP_ROW_KW)


def _is_total_row(name: str) -> bool:
    nl = name.lower().strip()
    return any(kw in nl for kw in _TOTAL_KW)


def _is_profit_row(name: str) -> bool:
    nl = name.lower().strip()
    return any(kw in nl for kw in _PROFIT_KW)


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

    sides = []  # (name_cols, amt_cols, positive_to_debit, indented)
    def _indented(name_cols):
        # Names occupy more than one column in the body → group is the outer
        # column, leaves are indented. Otherwise names are 'flat' (single
        # column) and header-vs-leaf is decided by amount presence.
        return sum(1 for c in name_cols if text_ct[c] > 0) > 1
    # Left side: liabilities/capital → Credit on a BS, expenses → Debit on a P&L
    sides.append((name_clusters[0], amt_clusters[0], is_pl, _indented(name_clusters[0])))
    # Right side (assets → Debit on a BS, income → Credit on a P&L)
    if len(name_clusters) >= 2 and len(amt_clusters) >= 2:
        sides.append((name_clusters[1], amt_clusters[1], not is_pl, _indented(name_clusters[1])))

    entries: List[TrialBalanceEntry] = []
    # Track group header totals so a group with detail rows does not also emit
    # its header total (which would double count).
    group_total = {}    # group name -> (amount, positive_to_debit)
    group_leaves = {}   # group name -> count of detail entries

    def _emit(name, group, amt, positive_to_debit):
        e = TrialBalanceEntry(account_name=name, group=group)
        if positive_to_debit:
            e.debit, e.credit = (abs(amt), 0.0) if amt >= 0 else (0.0, abs(amt))
        else:
            e.credit, e.debit = (abs(amt), 0.0) if amt >= 0 else (0.0, abs(amt))
        entries.append(e)

    def _rightmost_amt(row, cols):
        val = 0.0
        for c in cols:
            if 0 <= c < len(row):
                a = parse_amount(str(row[c]).strip())
                if a != 0.0:
                    val = a
        return val

    cur_group = ['' for _ in sides]
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

            if indented:
                # Group header in the OUTERMOST name column; sub-ledgers indented.
                outer = min(name_cols)
                group_txt = next((t for c, t in present if c == outer), None)
                leaf_txt  = next((t for c, t in present if c > outer), None)
                if leaf_txt and not _is_skip(leaf_txt):
                    if group_txt and not _is_skip(group_txt):
                        cur_group[si] = group_txt
                    g = cur_group[si] or group_txt or leaf_txt
                    if amt != 0.0:
                        _emit(leaf_txt, g, amt, ptd)
                        group_leaves[g] = group_leaves.get(g, 0) + 1
                elif group_txt and not _is_skip(group_txt):
                    cur_group[si] = group_txt
                    if amt != 0.0:
                        group_total[group_txt] = (amt, ptd)
            else:
                # FLAT layout (names share one column): a row with a name but NO
                # amount is a group header; a name WITH an amount is a leaf under
                # the most recent header.
                name = present[0][1]
                if _is_skip(name):
                    continue
                if amt == 0.0:
                    cur_group[si] = name           # group header line
                else:
                    g = cur_group[si] or name
                    _emit(name, g, amt, ptd)
                    group_leaves[g] = group_leaves.get(g, 0) + 1

    # Emit group-only totals for groups that produced no detail entries.
    for g, (amt, ptd) in group_total.items():
        if group_leaves.get(g, 0) == 0:
            _emit(g, g, amt, ptd)

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

    # ── Pass 1: collect raw rows ──────────────────────────────────────────────
    raw: List[Tuple[int, str, float, float]] = []  # (orig_row_idx, name, dr, cr)
    for row_idx in range(hdr_row + 1, len(df)):
        row = list(df.iloc[row_idx])
        while len(row) < 3:
            row.append('')

        name_raw = str(row[0]).strip()
        if name_raw.lower() in ('nan', ''):
            continue
        name = _clean(name_raw)
        if not name:
            continue
        if re.match(r'^\d{1,2}[-/][A-Za-z]', name) or re.match(r'^\d{4}$', name):
            continue
        if _is_total_row(name):
            continue

        dr = parse_amount(str(row[1]).strip() if len(row) > 1 else '')
        cr = parse_amount(str(row[2]).strip() if len(row) > 2 else '')
        raw.append((row_idx, name, dr, cr))

    if not raw:
        return []

    # ── Pass 2: classify parent vs leaf ───────────────────────────────────────
    n = len(raw)
    is_parent = [False] * n

    for i in range(n):
        _, _, dr_i, cr_i = raw[i]
        # Signal A: both cols non-zero → definitely a parent row
        if dr_i != 0.0 and cr_i != 0.0:
            is_parent[i] = True
            continue
        # Signal B: single-col row whose amount ≈ sum of following rows
        parent_amt = dr_i if dr_i != 0.0 else cr_i
        if parent_amt == 0.0:
            continue
        running = 0.0
        for j in range(i + 1, min(i + 80, n)):
            _, _, dr_j, cr_j = raw[j]
            # Stop at next signal-A row
            if dr_j != 0.0 and cr_j != 0.0:
                break
            child_amt = dr_j if dr_j != 0.0 else cr_j
            running += child_amt
            tol = max(1.0, parent_amt * 0.002)
            if abs(running - parent_amt) <= tol:
                is_parent[i] = True
                break
            if running > parent_amt * 1.02 + 500:
                break

    # ── Pass 3: emit entries ──────────────────────────────────────────────────
    entries: List[TrialBalanceEntry] = []
    current_group = ''
    # Track signal-A parents for fallback (if no leaves parsed under them)
    group_totals: Dict[str, Tuple[float, float]] = {}
    group_counts: Dict[str, int] = {}

    for i, (_, name, dr, cr) in enumerate(raw):
        if is_parent[i]:
            current_group = name
            _, _, p_dr, p_cr = raw[i]
            group_totals[name] = (p_dr, p_cr)
            group_counts.setdefault(name, 0)
            continue

        # Leaf row
        if dr == 0.0 and cr == 0.0:
            # No amount — treat as sub-group label
            current_group = name
            group_counts.setdefault(name, 0)
            continue

        e = TrialBalanceEntry(account_name=name, group=current_group)
        if dr != 0.0 and cr == 0.0:
            e.debit = dr
        elif cr != 0.0 and dr == 0.0:
            e.credit = cr
        else:
            net = dr - cr
            if net >= 0:
                e.debit = net
            else:
                e.credit = abs(net)

        entries.append(e)
        if current_group:
            group_counts[current_group] = group_counts.get(current_group, 0) + 1

    # ── Fallback: groups with no leaf entries → emit group total ──────────────
    for grp, (p_dr, p_cr) in group_totals.items():
        if group_counts.get(grp, 0) == 0:
            e = TrialBalanceEntry(account_name=grp, group=grp)
            if p_cr >= p_dr:
                e.credit = p_cr
            else:
                e.debit = p_dr
            entries.append(e)
            logger.info("Format I group-total fallback: %s Dr=%.2f Cr=%.2f",
                        grp, p_dr, p_cr)

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
    'total', 'grand total', 'page', 'from', 'to', 'account', 'amount',
    'liability', 'asset', 'liabilities', 'assets', 'carried', 'over',
    'brought', 'forward', 'balance sheet', 'profit and loss', 'profit & loss',
}
_J_SECTION_KW = {
    # BS Left (Credit / Liability)
    'capital account': 'CAPITAL',
    'partners capital': 'CAPITAL',
    'current liabilities': 'PROVISIONS',
    'duties and taxes': 'DUTIES AND TAXES',
    'brokers (sale commission agents)': 'SUNDRY CREDITORS',
    'provision': 'PROVISIONS',
    'sundry creditors': 'SUNDRY CREDITORS',
    'sundry creditors mill': 'SUNDRY CREDITORS',
    'sundry payables': 'SUNDRY CREDITORS',
    'loans liabilities': 'UNSECURED LOANS',
    'loans & borrowings': 'UNSECURED LOANS',
    # BS Right (Debit / Asset)
    'current assets': 'OTHER CURRENT ASSETS',
    'cash & bank balances': 'CASH AND BANK',
    'cash and bank balances': 'CASH AND BANK',
    'loans / advances a/c': 'LOANS AND ADVANCES (ASSETS)',
    'loans and advances': 'LOANS AND ADVANCES (ASSETS)',
    'sundry debtors': 'SUNDRY DEBTORS',
    'sundry receivables': 'SUNDRY DEBTORS',
    'fixed assets': 'FIXED ASSETS',
    'other current assets': 'OTHER CURRENT ASSETS',
    'stock in hand': 'CLOSING STOCK',
    'investments': 'INVESTMENTS',
    # P&L Left (Debit / Expenses)
    'opening stock': 'OPENING STOCK',
    'expenses direct': 'DIRECT EXPENSES (M)',
    'expenses direct (t&m)': 'DIRECT EXPENSES (M)',
    'purchase': 'PURCHASE A/C',
    'expenses indirect': 'INDIRECT EXPENSES',
    'expenses indirect (p&l)': 'INDIRECT EXPENSES',
    # P&L Right (Credit / Income)
    'sales': 'SALES A/C',
    'closing stock': 'CLOSING STOCK',
    'gross profit': 'CAPITAL',
    'other income': 'INDIRECT INCOMES',
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
            # Check for P&L (Account Name Amount repeated)
            is_pl = ('profit and loss' in text_lower or 'profit & loss' in text_lower
                     or 'account name amount' in text_lower.replace('\n', ' '))

            if not ((has_liab and has_asset) or is_pl):
                return None

            # Must NOT have a "Particulars" / "Debit" / "Credit" header
            # (those are Format H and Format I territory)
            if 'particulars' in text_lower:
                return None
            if 'debit' in text_lower and 'credit' in text_lower:
                return None

            # Find page split x by locating the right-side section header
            page_w = float(p0.width)

            # Find "Asset" or second "Account" word x position
            right_headers = [w for w in words
                             if w['text'].lower() in ('asset', 'assets', 'income')
                             and float(w['x0']) > page_w * 0.35]
            split_x = float(right_headers[0]['x0']) - 10 if right_headers else page_w * 0.48

            # Amount columns: find rightmost numeric word on a data row
            amt_words = [w for w in words
                         if re.match(r'^-?[\d,]+\.\d+$', w['text'].replace(',', '')
                                     .lstrip('-').lstrip('(').rstrip(')'))]
            left_amts  = sorted([float(w['x0']) for w in amt_words if float(w['x0']) < split_x], reverse=True)
            right_amts = sorted([float(w['x0']) for w in amt_words if float(w['x0']) >= split_x])
            left_amt_x  = left_amts[0]  if left_amts  else split_x * 0.85
            right_amt_x = right_amts[-1] if right_amts else page_w * 0.90

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

    _NUM_PAT = re.compile(r'^-?[\d,]+\.?\d*$')

    def _is_num(s: str) -> bool:
        return bool(_NUM_PAT.match(s.replace(',', '').lstrip('-').lstrip('(').rstrip(')')))

    def _parse_num(s: str) -> float:
        neg = s.strip().startswith('-') or (s.strip().startswith('(') and s.strip().endswith(')'))
        cleaned = s.replace(',', '').lstrip('-').lstrip('(').rstrip(')')
        try:
            val = float(cleaned)
            return -val if neg else val
        except ValueError:
            return 0.0

    def _emit(name: str, group: str, amt: float, is_dr: bool) -> None:
        if not name or abs(amt) < 0.01:
            return
        nl = name.lower().strip()
        if any(sk in nl for sk in _J_SKIP_KW):
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

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(keep_blank_chars=False)
            if not words:
                continue

            # Cluster into rows with tighter tolerance (2pt) to avoid merging
            # adjacent left/right rows that have slightly different y-positions
            words.sort(key=lambda w: (round(float(w['top']) / 2) * 2, float(w['x0'])))
            rows: Dict[int, list] = {}
            for w in words:
                y = round(float(w['top']) / 2) * 2
                rows.setdefault(y, []).append(w)

            for y in sorted(rows.keys()):
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

                    # Skip header/footer rows
                    skip_prefixes = ('account name', 'account amount', 'liability',
                                     'asset', 'amount', 'jai kanhaiya', '5/421',
                                     'ichalkaranji', '0230-', 'from :', 'from:',
                                     'balance sheet', 'profit and loss', 'profit & loss',
                                     'opening stock 47', 'expenses direct', 'purchase 21',
                                     'gross profit 22', 'grand total', 'expenses indirect',
                                     'net profit 1', 'sundry creditors write')
                    if any(nl.startswith(sk) for sk in skip_prefixes):
                        continue
                    if not name_str:
                        continue

                    cur_grp = left_group if side_name == 'L' else right_group

                    if not amt_tokens:
                        # No amount on this row — section header or pending name
                        if not any(sk in nl for sk in _J_SKIP_KW) \
                                and not _is_total_row(name_str):
                            canonical = _resolve_section(name_str, side_name)
                            if side_name == 'L':
                                left_group = canonical
                                left_pending_name = name_str
                            else:
                                right_group = canonical
                                right_pending_name = name_str
                    else:
                        # Has amounts
                        if side_name == 'L':
                            best = min(amt_tokens, key=lambda w: abs(float(w['x0']) - left_amt_x))
                        else:
                            best = min(amt_tokens, key=lambda w: abs(float(w['x0']) - right_amt_x))
                        amt = _parse_num(best['text'])

                        if name_str and not any(sk in nl for sk in _J_SKIP_KW) \
                                and not _is_total_row(name_str) \
                                and not re.match(r'^page\s+\d', nl):
                            _emit(name_str, cur_grp, amt, is_dr)
                            if side_name == 'L':
                                left_pending_name = ''
                            else:
                                right_pending_name = ''
                        else:
                            # Amount-only row — attach to pending name
                            pend = left_pending_name if side_name == 'L' else right_pending_name
                            grp  = left_group if side_name == 'L' else right_group
                            if pend:
                                _emit(pend, grp, amt, is_dr)
                                if side_name == 'L':
                                    left_pending_name = ''
                                else:
                                    right_pending_name = ''

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
    order = 0

    _SKIP_TEXT = {
        'debit', 'credit', 'closing', 'balance', 'particulars',
        'page', 'carried', 'over', 'continued', 'brought', 'forward',
        'trial', 'grand', 'total',
    }

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
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
                # Skip header/footer rows
                if any(sk in nl for sk in _SKIP_TEXT):
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
                order += 1

    if not raw:
        return []

    # ── Pass 2: detect parents (same logic as XLS Format I) ──────────────────
    n = len(raw)
    is_parent = [False] * n

    for i in range(n):
        _, _, dr_i, cr_i = raw[i]
        # Signal A: both non-zero → parent
        if dr_i != 0.0 and cr_i != 0.0:
            is_parent[i] = True
            continue
        # Signal B: sum-match lookahead
        parent_amt = dr_i if dr_i != 0.0 else cr_i
        if parent_amt == 0.0:
            continue
        running = 0.0
        for j in range(i + 1, min(i + 80, n)):
            _, _, dr_j, cr_j = raw[j]
            if dr_j != 0.0 and cr_j != 0.0:
                break
            child_amt = dr_j if dr_j != 0.0 else cr_j
            running += child_amt
            tol = max(1.0, parent_amt * 0.002)
            if abs(running - parent_amt) <= tol:
                is_parent[i] = True
                break
            if running > parent_amt * 1.02 + 500:
                break

    # ── Pass 3: emit leaf entries ─────────────────────────────────────────────
    entries: List[TrialBalanceEntry] = []
    current_group = ''
    group_totals: Dict[str, Tuple[float, float]] = {}
    group_counts: Dict[str, int] = {}

    for i, (_, name, dr, cr) in enumerate(raw):
        if is_parent[i]:
            current_group = name
            group_totals[name] = (dr, cr)
            group_counts.setdefault(name, 0)
            continue

        if dr == 0.0 and cr == 0.0:
            current_group = name
            group_counts.setdefault(name, 0)
            continue

        e = TrialBalanceEntry(account_name=name, group=current_group)
        if dr != 0.0 and cr == 0.0:
            e.debit = dr
        elif cr != 0.0 and dr == 0.0:
            e.credit = cr
        else:
            net = dr - cr
            if net >= 0:
                e.debit = net
            else:
                e.credit = abs(net)
        entries.append(e)
        if current_group:
            group_counts[current_group] = group_counts.get(current_group, 0) + 1

    # Fallback for groups with no leaf entries
    for grp, (p_dr, p_cr) in group_totals.items():
        if group_counts.get(grp, 0) == 0:
            e = TrialBalanceEntry(account_name=grp, group=grp)
            if p_cr >= p_dr:
                e.credit = p_cr
            else:
                e.debit = p_dr
            entries.append(e)
            logger.info("Format I-PDF group fallback: %s Dr=%.2f Cr=%.2f", grp, p_dr, p_cr)

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
