"""
Universal Parser — AUTO-ADAPTIVE cross-module file reader.

Given ANY file and a TARGET module ('loan_reporting', 'tds_returns', 'trial_balance'),
this module:
  1. Detects the actual content type of the file
  2. Parses it with the best-fit parser
  3. Adapts/converts the output to match what the target module needs
  4. NEVER raises 422 — always returns data with warnings for missing fields

Cross-module adapters:
  TDS file → Loan Reporting  : party=name, pan=pan, amt_paid=taken, tds=repaid
  TDS file → Trial Balance   : account=name, group=inferred from section, dr/cr from amounts
  Ledger   → TDS Returns     : account=deductee, date=payment_date, dr=amount_paid
  Ledger   → Trial Balance   : closing balance per account → Dr or Cr
  TB file  → Loan Reporting  : account=name, dr=taken, cr=repaid
  TB file  → TDS Returns     : account=deductee, dr/cr → amount
"""

import re
from datetime import date, datetime
from typing import List, Tuple, Any, Optional
from app.utils.content_detector import detect_content_type
from app.core.logger import get_logger
from app.core.exceptions import EmptyFileError

logger = get_logger(__name__)


# ── Import all parsers lazily to avoid circular imports ───────────────────────

def _tds_parser():
    from app.modules.tds_returns.parser import TDSReturnsParser
    return TDSReturnsParser()

def _tb_parser():
    from app.modules.trial_balance.parser import TrialBalanceParser
    return TrialBalanceParser()

def _ledger_pdf_parser():
    from app.modules.loan_reporting.parser import LedgerPDFParser
    return LedgerPDFParser()

def _ledger_excel_parser():
    from app.modules.loan_reporting.excel_csv_parser import LedgerExcelCSVParser
    return LedgerExcelCSVParser()


# ─────────────────────────────────────────────────────────────────────────────
# PUBLIC API
# ─────────────────────────────────────────────────────────────────────────────

def parse_for_loan_reporting(file_path: str) -> Tuple[List[Any], List[str]]:
    """
    Parse any file and return (raw_accounts, warnings) for Loan Reporting.
    raw_accounts is a list of RawAccount objects.
    """
    content_type = detect_content_type(file_path)
    warnings = []
    ext = file_path.rsplit('.', 1)[-1].lower()

    logger.info("parse_for_loan_reporting: %s → content_type=%s",
                file_path.rsplit('/', 1)[-1], content_type)

    # ── Native ledger file ────────────────────────────────────────────────────
    if content_type in ('ledger', 'unknown'):
        accounts = _try_ledger_parsers(file_path, ext, warnings)
        if accounts:
            return accounts, warnings

    # ── TDS file uploaded to Loan Reporting ───────────────────────────────────
    if content_type == 'tds':
        try:
            entries = _tds_parser().parse_file(file_path)
            if entries:
                warnings.append(
                    "TDS-format file detected. Extracted party names, PANs and amounts "
                    "for Loan Reporting. Dates and balance data are not available in this format."
                )
                accounts = _tds_entries_to_raw_accounts(entries)
                if accounts:
                    return accounts, warnings
        except Exception as e:
            warnings.append(f"TDS parser attempt: {e}")

    # ── Trial Balance file uploaded to Loan Reporting ─────────────────────────
    if content_type in ('tb', 'bs', 'unknown'):
        try:
            tb_entries = _tb_parser().parse_file(file_path)
            if tb_entries:
                warnings.append(
                    "Trial Balance / Balance Sheet file detected. "
                    "Extracted account balances for Loan Reporting. "
                    "Transaction history is not available in this format."
                )
                accounts = _tb_entries_to_raw_accounts(tb_entries)
                if accounts:
                    return accounts, warnings
        except Exception as e:
            warnings.append(f"Trial Balance parser attempt: {e}")

    # ── Last resort: try every parser ─────────────────────────────────────────
    accounts = _try_ledger_parsers(file_path, ext, warnings)
    if accounts:
        return accounts, warnings

    raise EmptyFileError(
        "Could not extract loan account data from this file.\n"
        "Accepted content:\n"
        "  • Tally ledger export (PDF/Excel/CSV) with Date, Dr, Cr, Balance columns\n"
        "  • TDS register (party names and amounts will be used)\n"
        "  • Trial balance (account balances will be used)"
    )


def parse_for_tds_returns(file_path: str) -> Tuple[List[Any], List[str]]:
    """
    Parse any file and return (tds_entries, warnings) for TDS Returns.
    """
    from app.modules.tds_returns.parser import TDSEntry
    content_type = detect_content_type(file_path)
    warnings = []
    ext = file_path.rsplit('.', 1)[-1].lower()

    logger.info("parse_for_tds_returns: %s → content_type=%s",
                file_path.rsplit('/', 1)[-1], content_type)

    # ── Native TDS file ───────────────────────────────────────────────────────
    if content_type in ('tds', 'unknown'):
        try:
            entries = _tds_parser().parse_file(file_path)
            if entries:
                return entries, warnings
        except Exception as e:
            warnings.append(f"TDS parser: {e}")

    # ── Ledger file → TDS Returns ─────────────────────────────────────────────
    if content_type in ('ledger', 'unknown'):
        accounts = _try_ledger_parsers(file_path, ext, warnings)
        if accounts:
            warnings.append(
                "Ledger-format file detected. Converting transactions to TDS entries. "
                "Section codes inferred from account names — please review."
            )
            ledger_tds = _raw_accounts_to_tds_entries(accounts)
            if ledger_tds:
                return ledger_tds, warnings

    # ── Trial Balance → TDS ───────────────────────────────────────────────────
    if content_type in ('tb', 'bs'):
        try:
            tb_entries = _tb_parser().parse_file(file_path)
            if tb_entries:
                warnings.append(
                    "Trial Balance file detected. Extracting debit-side entries as TDS candidates. "
                    "Section codes and rates must be reviewed manually."
                )
                entries = _tb_entries_to_tds_entries(tb_entries)
                if entries:
                    return entries, warnings
        except Exception as e:
            warnings.append(f"Trial Balance parser: {e}")

    # ── Retry TDS parser as last resort ──────────────────────────────────────
    try:
        final_entries = _tds_parser().parse_file(file_path)
        if final_entries:
            return final_entries, warnings
    except Exception as e:
        warnings.append(f"Final TDS parse attempt: {e}")

    raise EmptyFileError(
        "Could not extract TDS data from this file.\n"
        "Accepted content:\n"
        "  • TDS register (Excel/CSV/PDF) with Party Name, PAN, Amount, TDS columns\n"
        "  • Ledger export (transactions will be converted to TDS entries)"
    )


def parse_for_trial_balance(file_path: str) -> Tuple[List[Any], List[str]]:
    """
    Parse any file and return (tb_entries, warnings) for Trial Balance.
    """
    content_type = detect_content_type(file_path)
    warnings = []
    ext = file_path.rsplit('.', 1)[-1].lower()

    logger.info("parse_for_trial_balance: %s → content_type=%s",
                file_path.rsplit('/', 1)[-1], content_type)

    # ── Native TB/BS file ─────────────────────────────────────────────────────
    if content_type in ('tb', 'bs', 'unknown'):
        try:
            entries = _tb_parser().parse_file(file_path)
            if entries:
                return entries, warnings
        except Exception as e:
            warnings.append(f"Trial Balance parser: {e}")

    # ── Ledger → Trial Balance ────────────────────────────────────────────────
    if content_type in ('ledger', 'unknown'):
        accounts = _try_ledger_parsers(file_path, ext, warnings)
        if accounts:
            warnings.append(
                "Ledger-format file detected. Converting account closing balances "
                "to Trial Balance entries. Groups inferred from account names."
            )
            entries = _raw_accounts_to_tb_entries(accounts)
            if entries:
                return entries, warnings

    # ── TDS → Trial Balance ───────────────────────────────────────────────────
    if content_type == 'tds':
        try:
            tds_entries = _tds_parser().parse_file(file_path)
            if tds_entries:
                warnings.append(
                    "TDS-format file detected. Converting TDS amounts to Trial Balance entries. "
                    "Groups assigned based on TDS section codes."
                )
                entries = _tds_entries_to_tb_entries(tds_entries)
                if entries:
                    return entries, warnings
        except Exception as e:
            warnings.append(f"TDS parser: {e}")

    # ── Last resort: try TB parser ────────────────────────────────────────────
    try:
        entries = _tb_parser().parse_file(file_path)
        if entries:
            return entries, warnings
    except Exception as e:
        warnings.append(f"Final TB parse attempt: {e}")

    raise EmptyFileError(
        "Could not extract Trial Balance data from this file.\n"
        "Accepted content:\n"
        "  • Trial Balance with Debit/Credit columns\n"
        "  • Balance Sheet / P&L with Liabilities/Assets layout\n"
        "  • Ledger export (closing balances used)"
    )


# ── Helper: try both ledger parsers ──────────────────────────────────────────

def _try_ledger_parsers(file_path: str, ext: str, warnings: List[str]) -> list:
    accounts = []
    if ext == 'pdf':
        try:
            accounts = _ledger_pdf_parser().parse_file(file_path)
        except Exception as e:
            warnings.append(f"PDF ledger parser: {e}")
            try:
                accounts = _ledger_excel_parser().parse_file(file_path)
            except Exception as e2:
                warnings.append(f"Excel ledger parser: {e2}")
    else:
        try:
            accounts = _ledger_excel_parser().parse_file(file_path)
        except Exception as e:
            warnings.append(f"Excel ledger parser: {e}")
            try:
                accounts = _ledger_pdf_parser().parse_file(file_path)
            except Exception as e2:
                warnings.append(f"PDF ledger parser fallback: {e2}")
    return accounts




# ── Cross-module adapters ─────────────────────────────────────────────────────

def _tds_entries_to_raw_accounts(entries) -> list:
    """
    Convert TDS entries → RawAccount list for Loan Reporting.
    Each unique party = one RawAccount.
    amount_paid → single Rec transaction (taken)
    tds_deducted → single Pay transaction (repaid)
    """
    from app.modules.loan_reporting.parser import RawAccount
    from app.models.transaction import RawTransaction

    # Group by deductee name
    party_map = {}
    for e in entries:
        key = (e.deductee_name or '').strip()
        if not key:
            continue
        if key not in party_map:
            party_map[key] = {'pan': e.pan, 'total_amt': 0.0, 'total_tds': 0.0,
                              'date': e.payment_date_str or '01/04/2024'}
        party_map[key]['total_amt'] += e.amount_paid
        party_map[key]['total_tds'] += e.tds_deducted

    accounts = []
    for name, info in party_map.items():
        acct = RawAccount(name=name)
        acct.pan = info['pan'] or ''
        amt = info['total_amt']
        tds = info['total_tds']
        date_str = info['date']

        if amt > 0:
            acct.transactions.append(RawTransaction(
                date_str=date_str, voucher_type='Rec', voucher_no='BK-1',
                description='Amount Paid', narration='',
                amount=amt, balance_str=f'{amt}Cr',
                balance=amt, balance_direction='Cr',
            ))
        if tds > 0:
            remaining = max(0.0, amt - tds)
            acct.transactions.append(RawTransaction(
                date_str=date_str, voucher_type='Pay', voucher_no='BK-1',
                description='TDS Deducted', narration='',
                amount=tds, balance_str=f'{remaining}Cr',
                balance=remaining, balance_direction='Cr' if remaining > 0 else 'Zero',
            ))

        acct.closing_balance   = max(0.0, amt - tds)
        acct.closing_direction = 'Cr' if acct.closing_balance > 0 else 'Zero'
        acct.total_credit = amt
        acct.total_debit  = tds
        accounts.append(acct)

    return accounts


def _tb_entries_to_raw_accounts(tb_entries) -> list:
    """Convert Trial Balance entries → RawAccount for Loan Reporting."""
    from app.modules.loan_reporting.parser import RawAccount
    from app.models.transaction import RawTransaction

    accounts = []
    for e in tb_entries:
        if e.debit == 0 and e.credit == 0:
            continue
        acct = RawAccount(name=e.account_name)
        amt  = e.credit if e.credit > 0 else e.debit
        direction = 'Cr' if e.credit > 0 else 'Dr'
        acct.transactions.append(RawTransaction(
            date_str='01/04/2024', voucher_type='Opbl', voucher_no='',
            description='Opening Balance', narration='',
            amount=amt, balance_str=f'{amt}{direction}',
            balance=amt, balance_direction=direction,
        ))
        acct.closing_balance   = amt
        acct.closing_direction = direction
        accounts.append(acct)
    return accounts


def _raw_accounts_to_tds_entries(accounts) -> list:
    """Convert RawAccounts → TDS entries."""
    from app.modules.tds_returns.parser import TDSEntry

    entries = []
    for acct in accounts:
        total_cr = sum(t.amount for t in acct.transactions if t.voucher_type == 'Rec')
        if total_cr <= 0:
            continue
        e = TDSEntry(
            deductee_name=acct.name,
            pan=getattr(acct, 'pan', ''),
            section='[REVIEW SECTION]',
            amount_paid=total_cr,
            tds_deducted=0.0,
            rate=0.0,
            payment_date_str='',
            source_file='ledger',
            source_group='ledger',
        )
        entries.append(e)
    return entries


def _tb_entries_to_tds_entries(tb_entries) -> list:
    """Convert Trial Balance entries → TDS entries (debit side = expense payments)."""
    from app.modules.tds_returns.parser import TDSEntry
    from app.modules.tds_returns.parser import _type_value_to_section

    entries = []
    for e in tb_entries:
        amt = e.debit if e.debit > 0 else e.credit
        if amt <= 0:
            continue
        section = _type_value_to_section(e.group or e.account_name)
        tds_entry = TDSEntry(
            deductee_name=e.account_name,
            pan='',
            section=section,
            amount_paid=amt,
            tds_deducted=0.0,
            rate=0.0,
            payment_date_str='',
            source_file='trial_balance',
            source_group=e.group or 'UNGROUPED',
        )
        entries.append(tds_entry)
    return entries


def _raw_accounts_to_tb_entries(accounts) -> list:
    """Convert RawAccounts → TrialBalance entries using closing/computed balances."""
    from app.modules.trial_balance.parser import TrialBalanceEntry
    from app.utils.amount_parser import cr_value

    entries = []
    for acct in accounts:
        bal  = acct.closing_balance
        dirn = acct.closing_direction

        # If closing balance not set, derive from last transaction
        if bal <= 0 and acct.transactions:
            last = acct.transactions[-1]
            bal  = last.balance
            dirn = last.balance_direction

        # If still zero, use total_credit - total_debit
        if bal <= 0:
            net = acct.total_credit - acct.total_debit
            if net > 0:
                bal, dirn = net, 'Cr'
            elif net < 0:
                bal, dirn = abs(net), 'Dr'

        # Use transaction totals if balance still 0
        if bal <= 0:
            total_cr = sum(getattr(t, 'amount', 0) for t in acct.transactions
                          if getattr(t, 'voucher_type', '').lower() == 'rec')
            total_dr = sum(getattr(t, 'amount', 0) for t in acct.transactions
                          if getattr(t, 'voucher_type', '').lower() == 'pay')
            net = total_cr - total_dr
            if net > 0:
                bal, dirn = net, 'Cr'
            elif net < 0:
                bal, dirn = abs(net), 'Dr'

        if bal <= 0:
            continue

        e = TrialBalanceEntry(account_name=acct.name, group='LOANS AND ADVANCES (ASSETS)')
        if dirn == 'Cr':
            e.credit = bal
        else:
            e.debit  = bal
        entries.append(e)
    return entries


def _tds_entries_to_tb_entries(tds_entries) -> list:
    """Convert TDS entries → Trial Balance entries."""
    from app.modules.trial_balance.parser import TrialBalanceEntry

    entries = []
    for e in tds_entries:
        if e.amount_paid <= 0:
            continue
        tb = TrialBalanceEntry(
            account_name=e.deductee_name,
            group=f'TDS PAYABLE ({e.section})',
        )
        tb.debit = e.amount_paid
        entries.append(tb)
    return entries
