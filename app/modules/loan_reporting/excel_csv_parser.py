"""
Excel / CSV parser for Loan Reporting — ROBUST VERSION.

Accepts any tabular file (xlsx, xls, xlsm, ods, csv, tsv, txt) that contains
a ledger-style table of loan/advance transactions.

AUTO-ADAPTIVE strategy:
  1. Scan every sheet for a row containing Date + Dr/Cr/Balance columns
  2. Detect account-name rows (non-date rows between transaction blocks)
  3. Extract RawAccount / RawTransaction objects identical to the PDF parser output
  4. Handle merged header rows, multi-line headers, Indian number formats

Supported layouts:
  A) Multi-account: one table, account name in a merged/section header row
     Date | Voucher No | Particulars | Debit | Credit | Balance
  B) Single-account per sheet: sheet name = account name
  C) Tally-exported Excel: two merged header rows, account blocks separated by totals
"""

import re
import os
from typing import List, Optional, Dict, Tuple
import pandas as pd

from app.models.transaction import RawTransaction
from app.modules.loan_reporting.parser import RawAccount
from app.utils.amount_parser import parse_amount, parse_balance, cr_value
from app.core.logger import get_logger
from app.core.exceptions import FileParseError, EmptyFileError

logger = get_logger(__name__)

# ── Column alias lists ────────────────────────────────────────────────────────
_NAME_ALIASES = [
    'account name', 'party name', 'ledger name', 'ledger', 'account',
    'particulars', 'description', 'name of party', 'party',
]
_DATE_ALIASES = [
    'date', 'txn date', 'transaction date', 'voucher date', 'trans date',
    'dt', 'dated',
]
_DEBIT_ALIASES = [
    'debit', 'dr amount', 'dr.', 'dr', 'debit amount', 'amount dr',
    'withdrawal', 'paid out', 'outflow',
]
_CREDIT_ALIASES = [
    'credit', 'cr amount', 'cr.', 'cr', 'credit amount', 'amount cr',
    'deposit', 'received in', 'inflow',
]
_BALANCE_ALIASES = [
    'balance', 'closing balance', 'running balance', 'bal', 'current balance',
]
_VOUCHER_ALIASES = [
    'voucher no.', 'voucher no', 'vch no.', 'vch no', 'voucher', 'ref no',
    'ref', 'chq no', 'cheque no',
]
_NARRATION_ALIASES = [
    'narration', 'remarks', 'particulars', 'description', 'details',
]

_DATE_PATTERNS = [
    re.compile(r'^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$'),    # DD/MM/YYYY or DD-MM-YY
    re.compile(r'^\d{4}[-/]\d{2}[-/]\d{2}$'),           # YYYY-MM-DD
    re.compile(r'^\d{1,2}\s+\w+\s+\d{4}$'),             # DD Mon YYYY
]

_AMOUNT_RE  = re.compile(r'^[\-\(]?[\d,]+\.?\d*[\)]?$')
_SKIP_WORDS = frozenset([
    'total', 'sub total', 'sub-total', 'grand total', 'subtotal',
    'closing balance', 'opening balance', 'balance b/d', 'balance b.d.',
    'brought forward', 'carried forward',
    'date', 'voucher', 'particulars', 'narration', 'debit', 'credit',
    'dr', 'cr', 'balance', 'amount', 'ledger', 'account',
])


def _is_date(val: str) -> bool:
    val = val.strip()
    if not val or val.lower() in ('nan', 'none', ''):
        return False
    for pat in _DATE_PATTERNS:
        if pat.match(val):
            return True
    # Also accept pandas Timestamp string representation
    if re.match(r'\d{4}-\d{2}-\d{2}\s', val):
        return True
    return False


def _normalise_date(val: str) -> str:
    """Convert any date format to DD/MM/YYYY."""
    val = val.strip()
    # YYYY-MM-DD (possibly with time)
    m = re.match(r'(\d{4})[-/](\d{2})[-/](\d{2})', val)
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    # DD/MM/YYYY or DD-MM-YYYY
    m = re.match(r'(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})', val)
    if m:
        d, mo, y = m.group(1), m.group(2), m.group(3)
        if len(y) == 2:
            y = '20' + y
        return f"{int(d):02d}/{int(mo):02d}/{y}"
    return val


def _sv(row, idx: int) -> str:
    if 0 <= idx < len(row):
        v = str(row[idx]).strip()
        return '' if v.lower() in ('nan', 'none', '') else v
    return ''


def _find_col(header_lower: List[str], aliases: List[str], exclude: set = None) -> int:
    """Find column index by alias matching. exclude = set of already-claimed indices."""
    exclude = exclude or set()
    # Pass 1: exact match or alias is substring of header
    for alias in aliases:
        for ci, h in enumerate(header_lower):
            if ci in exclude or not h or h in ('nan',):
                continue
            if alias == h or alias in h:
                return ci
    # Pass 2: header is substring of alias (weak)
    for alias in aliases:
        for ci, h in enumerate(header_lower):
            if ci in exclude or not h or h in ('nan',):
                continue
            if len(h) >= 2 and h in alias:
                return ci
    return -1


def _is_total_row(row: list) -> bool:
    """Return True if row looks like a totals/closing row."""
    for v in row[:4]:
        s = str(v).strip().lower()
        if s in ('nan', ''):
            continue
        if any(kw in s for kw in ('total', 'closing', 'balance b/', 'brought', 'carried')):
            return True
    return False


def _is_account_header(val: str) -> bool:
    """Return True if val looks like an account name header row."""
    v = val.strip()
    if len(v) < 3:
        return False
    if v.lower() in _SKIP_WORDS:
        return False
    if _is_date(v):
        return False
    if _AMOUNT_RE.match(v.replace(',', '')):
        return False
    if not re.search(r'[A-Za-z]{2,}', v):
        return False
    if re.match(r'^\d+$', v):
        return False
    return True


class LedgerExcelCSVParser:
    """
    Parses Tally ledger data from Excel/CSV files.
    """

    def parse_file(self, file_path: str) -> List[RawAccount]:
        ext = file_path.rsplit('.', 1)[-1].lower()
        all_dfs: Dict[str, pd.DataFrame] = {}

        try:
            if ext in ('xlsx', 'xlsm'):
                all_dfs = pd.read_excel(
                    file_path, sheet_name=None, header=None,
                    dtype=str, engine='openpyxl'
                )
            elif ext == 'xls':
                all_dfs = pd.read_excel(
                    file_path, sheet_name=None, header=None,
                    dtype=str, engine='xlrd'
                )
            elif ext == 'ods':
                all_dfs = pd.read_excel(
                    file_path, sheet_name=None, header=None,
                    dtype=str, engine='odf'
                )
            elif ext in ('csv', 'tsv', 'txt'):
                sep = '\t' if ext == 'tsv' else None
                df = pd.read_csv(
                    file_path, header=None, dtype=str,
                    sep=sep, engine='python', on_bad_lines='skip',
                    encoding_errors='replace'
                )
                all_dfs = {'Sheet1': df}
            else:
                # Try CSV as last resort
                df = pd.read_csv(
                    file_path, header=None, dtype=str,
                    engine='python', on_bad_lines='skip',
                    encoding_errors='replace'
                )
                all_dfs = {'Sheet1': df}
        except Exception as e:
            raise FileParseError(f"Could not read file '{os.path.basename(file_path)}': {e}")

        accounts: List[RawAccount] = []
        for sheet_name, df in all_dfs.items():
            df = df.fillna('').map(lambda x: str(x).strip())
            try:
                sheet_accounts = self._parse_sheet(df, sheet_name)
                accounts.extend(sheet_accounts)
            except Exception as e:
                logger.warning("Sheet '%s' parse error: %s", sheet_name, e)

        if not accounts:
            raise EmptyFileError(
                "No loan/advance transactions detected.\n"
                "Expected columns: Date, Debit/Dr, Credit/Cr, Balance (or Dr+Cr)\n"
                "and account name separators between transaction groups."
            )

        logger.info(
            "Excel/CSV parser: %d accounts from %s",
            len(accounts), os.path.basename(file_path)
        )
        return accounts

    # ── Sheet-level parsing ───────────────────────────────────────────────────

    def _parse_sheet(self, df: pd.DataFrame, sheet_name: str) -> List[RawAccount]:
        header_row, col_map = self._detect_header(df)
        if header_row < 0:
            logger.debug("Sheet '%s': no recognisable transaction header", sheet_name)
            return []
        return self._extract_accounts(df, header_row, col_map, sheet_name)

    def _detect_header(self, df: pd.DataFrame) -> Tuple[int, Dict[str, int]]:
        """
        Scan rows 0–40 for a header row containing at least:
          - date column
          - debit OR credit column (or a combined Dr/Cr amount column)
        """
        for ri in range(min(40, len(df))):
            row = [str(c).lower().strip() for c in df.iloc[ri]]

            claimed: set = set()

            col_date    = _find_col(row, _DATE_ALIASES, claimed)
            if col_date >= 0: claimed.add(col_date)

            col_debit   = _find_col(row, _DEBIT_ALIASES, claimed)
            if col_debit >= 0: claimed.add(col_debit)

            col_credit  = _find_col(row, _CREDIT_ALIASES, claimed)
            if col_credit >= 0: claimed.add(col_credit)

            col_balance = _find_col(row, _BALANCE_ALIASES, claimed)
            if col_balance >= 0: claimed.add(col_balance)

            col_account = _find_col(row, _NAME_ALIASES, claimed)
            if col_account >= 0: claimed.add(col_account)

            col_voucher = _find_col(row, _VOUCHER_ALIASES, claimed)
            if col_voucher >= 0: claimed.add(col_voucher)

            col_narr    = _find_col(row, _NARRATION_ALIASES, claimed)

            # Collision guard: debit == credit → clear credit
            if col_debit >= 0 and col_debit == col_credit:
                col_credit = -1

            # Must have date + at least one money column
            has_date   = col_date >= 0
            has_money  = col_debit >= 0 or col_credit >= 0 or col_balance >= 0

            # Also validate: check that the next row actually has a date-like value
            # in the date column (confirms this is really a header, not a data row)
            if has_date and has_money:
                # Quick confirm: at least one of the next 5 rows has a date in col_date
                confirmed = False
                for check_ri in range(ri + 1, min(ri + 6, len(df))):
                    cell_val = str(df.iloc[check_ri, col_date]).strip()
                    if _is_date(cell_val):
                        confirmed = True
                        break
                if confirmed:
                    col_map = {
                        'date': col_date,
                        'debit': col_debit,
                        'credit': col_credit,
                        'balance': col_balance,
                        'account': col_account,
                        'voucher': col_voucher,
                        'narration': col_narr,
                        'header_row': ri,
                    }
                    logger.debug(
                        "Header at row %d: date=%d dr=%d cr=%d bal=%d acct=%d",
                        ri, col_date, col_debit, col_credit, col_balance, col_account
                    )
                    return ri, col_map

        return -1, {}

    # ── Account extraction ────────────────────────────────────────────────────

    def _extract_accounts(
        self,
        df: pd.DataFrame,
        header_row: int,
        col_map: Dict[str, int],
        sheet_name: str
    ) -> List[RawAccount]:

        col_date    = col_map['date']
        col_debit   = col_map['debit']
        col_credit  = col_map['credit']
        col_balance = col_map['balance']
        col_account = col_map['account']
        col_voucher = col_map['voucher']
        col_narr    = col_map['narration']

        accounts: List[RawAccount] = []
        current: Optional[RawAccount] = None
        prev_cr: float = 0.0

        # Use sheet name as default account name if no headers found
        default_name = sheet_name.strip() if sheet_name.strip() else 'Account'

        for ri in range(header_row + 1, len(df)):
            row = list(df.iloc[ri])

            date_val = _sv(row, col_date) if col_date >= 0 else ''

            # ── Total / closing row — capture totals, close account ─────────
            if _is_total_row(row):
                if current and current.transactions:
                    # Try to capture closing balance from balance column
                    if col_balance >= 0:
                        bal_raw = _sv(row, col_balance)
                        if bal_raw:
                            amt, direction = parse_balance(bal_raw)
                            if amt:
                                current.closing_balance   = amt
                                current.closing_direction = direction
                    # Capture totals from Dr/Cr columns
                    if col_debit >= 0:
                        dr_tot = parse_amount(_sv(row, col_debit))
                        if dr_tot: current.total_debit = dr_tot
                    if col_credit >= 0:
                        cr_tot = parse_amount(_sv(row, col_credit))
                        if cr_tot: current.total_credit = cr_tot

                    accounts.append(current)
                    current = None
                    prev_cr = 0.0
                continue

            # ── Transaction row (starts with a date) ─────────────────────────
            if _is_date(date_val):
                if current is None:
                    # Auto-create account from sheet name
                    current = RawAccount(name=default_name)
                    prev_cr = 0.0

                date_norm = _normalise_date(date_val)

                debit_raw  = _sv(row, col_debit)   if col_debit  >= 0 else ''
                credit_raw = _sv(row, col_credit)  if col_credit >= 0 else ''
                bal_raw    = _sv(row, col_balance) if col_balance >= 0 else ''
                voucher_no = _sv(row, col_voucher) if col_voucher >= 0 else ''
                narration  = _sv(row, col_narr)    if col_narr   >= 0 else ''

                # Particulars / description (from account col if it's a text desc column)
                description = _sv(row, col_account) if col_account >= 0 else ''

                debit_amt  = parse_amount(debit_raw)
                credit_amt = parse_amount(credit_raw)

                # Some formats put the amount in a single column with Cr/Dr suffix
                # e.g. col_debit contains "12,345.00 Dr" or "12,345.00 Cr"
                if debit_amt == 0 and credit_amt == 0 and col_debit >= 0:
                    raw_combined = debit_raw or credit_raw
                    if raw_combined:
                        amt_val, direction = parse_balance(raw_combined)
                        if amt_val:
                            if direction == 'Dr':
                                debit_amt = amt_val
                            elif direction == 'Cr':
                                credit_amt = amt_val

                amount = debit_amt if debit_amt else credit_amt

                # Parse running balance
                balance     = 0.0
                bal_dir     = 'Zero'
                if bal_raw:
                    balance, bal_dir = parse_balance(bal_raw)
                elif amount:
                    # Infer from previous balance + this transaction
                    new_cr = prev_cr + credit_amt - debit_amt
                    balance = abs(new_cr)
                    bal_dir = 'Cr' if new_cr >= 0 else ('Dr' if new_cr < 0 else 'Zero')

                cur_cr = cr_value(balance, bal_dir)

                # Determine voucher type
                desc_lower = (description + ' ' + narration).lower()
                if re.search(r'opening\s+balance|balance\s+b[/.]?d', desc_lower):
                    v_type = 'Opbl'
                    voucher_no = voucher_no or ''
                elif debit_amt > 0 and credit_amt == 0:
                    v_type = 'Pay'
                    voucher_no = voucher_no or 'BK-1'
                elif credit_amt > 0 and debit_amt == 0:
                    v_type = 'Rec'
                    voucher_no = voucher_no or 'BK-1'
                else:
                    # Both or neither — infer from balance movement
                    delta = cur_cr - prev_cr
                    v_type = 'Rec' if delta >= 0 else 'Pay'
                    voucher_no = voucher_no or 'BK-1'

                # Tag interest/TDS as journal entries
                if re.search(r'\binterest\b|\btds\b|t\.d\.s', desc_lower):
                    v_type = 'Jrn'
                    voucher_no = ''

                txn = RawTransaction(
                    date_str=date_norm,
                    voucher_type=v_type,
                    voucher_no=voucher_no,
                    description=description or narration or '',
                    narration=narration,
                    amount=amount,
                    balance_str=bal_raw or f"{balance}{bal_dir}",
                    balance=balance,
                    balance_direction=bal_dir,
                )
                current.transactions.append(txn)
                prev_cr = cur_cr
                continue

            # ── Non-date, non-total row — could be account name ───────────────
            name_candidate = ''

            # Check dedicated account column first
            if col_account >= 0:
                v = _sv(row, col_account)
                if v and _is_account_header(v):
                    name_candidate = v

            # If no account column, scan first 3 columns for a text name
            if not name_candidate:
                for ci in range(min(3, len(row))):
                    v = _sv(row, ci)
                    if v and _is_account_header(v) and not _is_date(v):
                        # Make sure this row has no amounts in debit/credit columns
                        has_amounts = False
                        for money_ci in [col_debit, col_credit]:
                            if money_ci >= 0:
                                mv = _sv(row, money_ci)
                                if mv and parse_amount(mv) != 0:
                                    has_amounts = True
                                    break
                        if not has_amounts:
                            name_candidate = v
                            break

            if name_candidate:
                if current and current.transactions:
                    accounts.append(current)
                current = RawAccount(name=name_candidate)
                prev_cr = 0.0
                logger.debug("New account from Excel: %s", name_candidate)

        # Close last account
        if current and current.transactions:
            accounts.append(current)

        return accounts
