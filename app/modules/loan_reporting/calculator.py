"""
Calculator for Loan Reporting (Form 3CD Clause 31 – Loans/Deposits).

Calculations per account:
  taken    = sum of all Rec (BK-prefix) credit transactions during the year
  repaid   = sum of all Pay (BK-prefix) debit transactions during the year
  maximum  = peak Cr balance treating same journal-number entries as atomic
  squared_up = YES if closing balance is zero

Filter rule:
  Include only accounts that had a Cr balance at any real checkpoint
  during the year (i.e., they represent deposits/loans received by MTA).
  Exclude purely Dr accounts (loans given by MTA, reported separately).
"""
from typing import List, Optional, Tuple
from app.models.account import NormalizedAccount, LoanReportRow
from app.models.transaction import NormalizedTransaction
from app.utils.amount_parser import cr_value
from app.core.logger import get_logger

logger = get_logger(__name__)


def _compute_taken(transactions: List[NormalizedTransaction]) -> float:
    """Sum of all bank Receipt (Rec + BK-prefix) credit amounts."""
    return sum(
        t.credit
        for t in transactions
        if t.voucher_type.lower() == 'rec' and t.is_bank_txn
    )


def _compute_repaid(transactions: List[NormalizedTransaction]) -> float:
    """Sum of all bank Payment (Pay + BK-prefix) debit amounts."""
    return sum(
        t.debit
        for t in transactions
        if t.voucher_type.lower() == 'pay' and t.is_bank_txn
    )


def _compute_maximum(
    opening_balance: float,
    opening_direction: str,
    transactions: List[NormalizedTransaction],
) -> float:
    """
    Maximum Cr balance treating same-journal-number entries as atomic.

    Algorithm:
      - Start with opening balance as first checkpoint
      - For bank transactions: each entry is a real checkpoint
      - For journal entries: group by journal_no; only the FINAL balance
        of the group counts as a checkpoint (intermediate sub-entries ignored)
      - Maximum Cr value across all checkpoints
    """
    checkpoints: List[float] = [cr_value(opening_balance, opening_direction)]

    i = 0
    non_opening = [t for t in transactions if not t.is_opening]
    while i < len(non_opening):
        txn = non_opening[i]

        if txn.is_journal and txn.journal_no:
            # Collect all entries of the same journal
            jrn_no = txn.journal_no
            j = i
            last_balance = txn.balance
            last_direction = txn.balance_direction
            while j < len(non_opening) and non_opening[j].journal_no == jrn_no:
                last_balance = non_opening[j].balance
                last_direction = non_opening[j].balance_direction
                j += 1
            checkpoints.append(cr_value(last_balance, last_direction))
            i = j
        else:
            checkpoints.append(cr_value(txn.balance, txn.balance_direction))
            i += 1

    max_cr = max(checkpoints) if checkpoints else 0.0
    return max(max_cr, 0.0)  # never negative


def _had_cr_balance(
    opening_balance: float,
    opening_direction: str,
    transactions: List[NormalizedTransaction],
) -> bool:
    """Return True if this account ever had a Cr balance during the year."""
    if opening_direction == 'Cr' and opening_balance > 0:
        return True
    for t in transactions:
        if t.balance_direction == 'Cr' and t.balance > 0:
            return True
    return False


class LoanReportingCalculator:

    def compute_report_rows(
        self,
        accounts: List[NormalizedAccount],
    ) -> List[LoanReportRow]:
        rows: List[LoanReportRow] = []
        sr = 1

        for acct in accounts:
            # Exclude purely Dr accounts (loans given by MTA, not received)
            if not _had_cr_balance(
                acct.opening_balance, acct.opening_direction, acct.transactions
            ):
                logger.debug("Excluding Dr-only account: %s", acct.name)
                continue

            # Exclude Dr-closing accounts: these changed from Cr → Dr during the year
            # (net position reversed — they are now debtors to MTA, not creditors)
            # Clause 31 reports only loans/deposits MTA received (Cr or settled to zero)
            if acct.closing_direction == 'Dr':
                logger.debug("Excluding Dr-closing account: %s", acct.name)
                continue

            taken = _compute_taken(acct.transactions)
            repaid = _compute_repaid(acct.transactions)
            maximum = _compute_maximum(
                acct.opening_balance, acct.opening_direction, acct.transactions
            )
            squared_up = (
                acct.closing_direction == 'Zero'
                or acct.closing_balance == 0.0
            )

            row = LoanReportRow(
                sr_no=sr,
                name=acct.name,
                address=acct.location,
                pan=acct.pan or '',
                taken=taken,
                repaid=repaid,
                maximum=maximum,
                squared_up='YES' if squared_up else 'NO',
            )
            rows.append(row)
            sr += 1
            logger.debug(
                "Account %r → taken=%.2f repaid=%.2f max=%.2f squared=%s",
                acct.name, taken, repaid, maximum, row.squared_up,
            )

        return rows
