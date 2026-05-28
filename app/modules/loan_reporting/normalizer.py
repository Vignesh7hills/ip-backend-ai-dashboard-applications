"""
Normalizer for Loan Reporting.

Converts RawAccount + RawTransaction → NormalizedAccount + NormalizedTransaction.

Key logic:
  • Determine debit/credit column from voucher type + balance change
  • Tag bank transactions (BK-prefix vouchers)
  • Tag journal entries and group by journal number
  • Resolve opening balance direction
"""
import re
from typing import List
from app.modules.loan_reporting.parser import RawAccount
from app.models.transaction import RawTransaction, NormalizedTransaction
from app.models.account import NormalizedAccount
from app.utils.amount_parser import parse_amount, parse_balance, cr_value
from app.utils.date_parser import parse_date
from app.core.logger import get_logger

logger = get_logger(__name__)

_BK_RE = re.compile(r'^BK-\d+$', re.IGNORECASE)


def _is_bank_txn(voucher_no: str) -> bool:
    return bool(_BK_RE.match(voucher_no.strip()))


def _resolve_debit_credit(
    raw: RawTransaction, prev_cr_value: float
) -> tuple[float, float]:
    """
    Determine the debit and credit amounts for a transaction.

    Strategy:
      - Pay (bank) → debit
      - Rec (bank) → credit
      - Opbl → opening (credit if Cr, debit if Dr, no debit/credit in the traditional sense)
      - Jrn → determine from balance change vs previous balance

    Returns (debit, credit) — one will always be 0.
    """
    v = raw.voucher_type.lower()
    amount = raw.amount

    if v == 'opbl':
        return 0.0, 0.0   # opening balance — handled separately

    current_cr = cr_value(raw.balance, raw.balance_direction)
    delta = current_cr - prev_cr_value

    if v == 'pay':
        # Payment: reduces Cr (or increases Dr) → debit side
        return amount, 0.0
    if v == 'rec':
        # Receipt: increases Cr (or reduces Dr) → credit side
        return 0.0, amount
    if v == 'jrn':
        # Journal: use balance delta to determine direction
        if delta >= 0:
            return 0.0, amount   # balance increased → credit
        else:
            return amount, 0.0   # balance decreased → debit
    # Fallback
    return amount, 0.0


class LoanReportingNormalizer:

    def normalize(self, raw_accounts: List[RawAccount]) -> List[NormalizedAccount]:
        result = []
        for raw in raw_accounts:
            try:
                norm = self._normalize_account(raw)
                result.append(norm)
            except Exception as exc:
                logger.warning("Skipping account %r due to normalization error: %s", raw.name, exc)
        return result

    def _normalize_account(self, raw: RawAccount) -> NormalizedAccount:
        # Determine opening balance from first Opbl transaction
        opening_balance = 0.0
        opening_direction = 'Zero'
        for txn in raw.transactions:
            if txn.voucher_type.lower() == 'opbl':
                opening_balance = txn.balance
                opening_direction = txn.balance_direction
                break

        # Normalize transactions with debit/credit resolution
        norm_txns: List[NormalizedTransaction] = []
        prev_cr = cr_value(opening_balance, opening_direction)

        for raw_txn in raw.transactions:
            debit, credit = _resolve_debit_credit(raw_txn, prev_cr)
            current_cr = cr_value(raw_txn.balance, raw_txn.balance_direction)

            norm = NormalizedTransaction(
                txn_date=parse_date(raw_txn.date_str),
                voucher_type=raw_txn.voucher_type,
                voucher_no=raw_txn.voucher_no,
                description=raw_txn.description,
                narration=raw_txn.narration,
                debit=debit,
                credit=credit,
                balance=raw_txn.balance,
                balance_direction=raw_txn.balance_direction,
                is_bank_txn=_is_bank_txn(raw_txn.voucher_no),
                is_journal=raw_txn.voucher_type.lower() == 'jrn',
                is_opening=raw_txn.voucher_type.lower() == 'opbl',
                journal_no=raw_txn.voucher_no if raw_txn.voucher_type.lower() == 'jrn' else '',
            )
            norm_txns.append(norm)
            if not norm.is_opening:
                prev_cr = current_cr

        acct = NormalizedAccount(
            name=raw.name,
            location=raw.location,
            pan=getattr(raw, 'pan', ''),
            opening_balance=opening_balance,
            opening_direction=opening_direction,
            closing_balance=raw.closing_balance,
            closing_direction=raw.closing_direction,
            total_debit=raw.total_debit,
            total_credit=raw.total_credit,
            transactions=norm_txns,
        )
        return acct
