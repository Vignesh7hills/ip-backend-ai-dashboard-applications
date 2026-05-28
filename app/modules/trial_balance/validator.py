"""
Trial Balance Validator.

Checks:
  1. Dr/Cr totals balance (fundamental double-entry requirement)
  2. No account has both Dr and Cr simultaneously
  3. No negative amounts
  4. Duplicate account name detection
"""
from typing import List, Tuple
from app.modules.trial_balance.parser import TrialBalanceEntry
from app.core.logger import get_logger

logger = get_logger(__name__)
_TOLERANCE = 1.0   # Rupee tolerance for totals mismatch


class TrialBalanceValidator:

    def validate(
        self, entries: List[TrialBalanceEntry]
    ) -> Tuple[bool, List[str], List[str]]:
        errors: List[str] = []
        warnings: List[str] = []
        data = [e for e in entries if not e.is_total and not e.is_subtotal]

        # ── 1. Dr = Cr totals ─────────────────────────────────────────────────
        total_dr = sum(e.debit for e in data)
        total_cr = sum(e.credit for e in data)
        diff = abs(total_dr - total_cr)
        if diff > _TOLERANCE:
            errors.append(
                f"Trial balance does not tally: Total Dr={total_dr:,.2f}, "
                f"Total Cr={total_cr:,.2f}, Difference={diff:,.2f}"
            )
        else:
            logger.info("Trial balance tallied: Dr=Cr=%.2f", total_dr)

        # ── 2. Both Dr and Cr on same account ─────────────────────────────────
        for e in data:
            if e.debit > 0 and e.credit > 0:
                warnings.append(
                    f"Account '{e.account_name}' has both Dr ({e.debit:,.2f}) "
                    f"and Cr ({e.credit:,.2f}) — verify correctness."
                )

        # ── 3. Negative amounts ───────────────────────────────────────────────
        for e in data:
            if e.debit < 0 or e.credit < 0:
                errors.append(
                    f"Negative amount in account '{e.account_name}': "
                    f"Dr={e.debit:,.2f}, Cr={e.credit:,.2f}"
                )

        # ── 4. Duplicates ─────────────────────────────────────────────────────
        seen = {}
        for e in data:
            name = e.account_name.strip().lower()
            if name in seen:
                warnings.append(f"Duplicate account name: '{e.account_name}'")
            seen[name] = True

        return len(errors) == 0, errors, warnings
