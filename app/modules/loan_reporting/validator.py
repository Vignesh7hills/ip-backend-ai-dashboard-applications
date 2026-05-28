"""
Validation engine for Loan Reporting.

Checks performed:
  1. Balance continuity: each transaction's balance must follow from prior
  2. Account totals: PDF-supplied totals vs computed totals
  3. Date order: transactions should be in ascending date order
  4. Duplicate transaction detection (same date + voucher_no + amount)
  5. Empty / malformed row detection
"""
from typing import List, Tuple
from app.models.account import NormalizedAccount
from app.models.transaction import NormalizedTransaction
from app.utils.amount_parser import cr_value
from app.core.logger import get_logger

logger = get_logger(__name__)

_TOLERANCE = 0.05   # float rounding tolerance in rupees


def _near(a: float, b: float) -> bool:
    return abs(a - b) <= _TOLERANCE


class LoanReportingValidator:

    def validate_all(
        self, accounts: List[NormalizedAccount]
    ) -> List[NormalizedAccount]:
        for acct in accounts:
            errors, warnings = self._validate_account(acct)
            acct.validation_errors = errors
            acct.balance_validated = len(errors) == 0
            if errors:
                logger.warning("Account %r has %d validation errors", acct.name, len(errors))
            if warnings:
                for w in warnings:
                    logger.debug("Account %r warning: %s", acct.name, w)
        return accounts

    def _validate_account(
        self, acct: NormalizedAccount
    ) -> Tuple[List[str], List[str]]:
        errors: List[str] = []
        warnings: List[str] = []
        txns = acct.transactions

        if not txns:
            warnings.append("No transactions found.")
            return errors, warnings

        # ── 1. Balance continuity ─────────────────────────────────────────────
        prev_cr = cr_value(acct.opening_balance, acct.opening_direction)
        for txn in txns:
            if txn.is_opening:
                continue
            expected_cr = prev_cr + txn.credit - txn.debit
            actual_cr = cr_value(txn.balance, txn.balance_direction)
            if not _near(expected_cr, actual_cr):
                errors.append(
                    f"Balance mismatch on {txn.txn_date} {txn.voucher_no}: "
                    f"expected {expected_cr:.2f}, got {actual_cr:.2f}"
                )
            prev_cr = actual_cr

        # ── 2. Closing balance matches PDF totals ─────────────────────────────
        if txns:
            last_non_opening = [t for t in txns if not t.is_opening]
            if last_non_opening:
                computed_closing = cr_value(
                    last_non_opening[-1].balance,
                    last_non_opening[-1].balance_direction,
                )
                pdf_closing = cr_value(acct.closing_balance, acct.closing_direction)
                if not _near(computed_closing, pdf_closing):
                    errors.append(
                        f"Closing balance mismatch: computed={computed_closing:.2f}, "
                        f"PDF supplied={pdf_closing:.2f}"
                    )

        # ── 3. Date order ─────────────────────────────────────────────────────
        dated = [t for t in txns if t.txn_date and not t.is_opening]
        for j in range(1, len(dated)):
            if dated[j].txn_date < dated[j - 1].txn_date:
                warnings.append(
                    f"Out-of-order date: {dated[j].txn_date} after {dated[j-1].txn_date}"
                )
                break  # report once per account

        # ── 4. Duplicate detection ────────────────────────────────────────────
        seen = set()
        for txn in txns:
            key = (txn.txn_date, txn.voucher_no, txn.debit, txn.credit)
            if key in seen and txn.voucher_no:
                warnings.append(
                    f"Possible duplicate: {txn.txn_date} {txn.voucher_no} "
                    f"Dr={txn.debit} Cr={txn.credit}"
                )
            seen.add(key)

        # ── 5. Malformed rows ─────────────────────────────────────────────────
        for txn in txns:
            if txn.debit == 0 and txn.credit == 0 and not txn.is_opening:
                warnings.append(
                    f"Zero-amount row on {txn.txn_date} {txn.voucher_no}: "
                    f"{txn.description}"
                )

        return errors, warnings
