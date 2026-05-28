"""TDS Returns Validator."""
from typing import List, Tuple
from app.modules.tds_returns.parser import TDSEntry
from app.core.logger import get_logger
import re

logger = get_logger(__name__)
_PAN_RE = re.compile(r'^[A-Z]{5}\d{4}[A-Z]$')
_KNOWN_SECTIONS = {
    '192', '192A', '193', '194', '194A', '194B', '194C', '194D',
    '194DA', '194E', '194EE', '194F', '194G', '194H', '194I',
    '194IA', '194IB', '194J', '194K', '194LA', '194LB', '194LC',
    '194LD', '194M', '194N', '194O', '194P', '194Q', '194R',
    '194S', '195', '196B', '196C', '196D',
}


class TDSReturnsValidator:

    def validate(
        self, entries: List[TDSEntry]
    ) -> Tuple[bool, List[str], List[str]]:
        errors: List[str] = []
        warnings: List[str] = []

        for i, e in enumerate(entries, 1):
            prefix = f"Row {i} [{e.deductee_name or 'Unknown'}]"

            # PAN validation
            if not e.pan:
                warnings.append(f"{prefix}: Missing PAN")
            elif not _PAN_RE.match(e.pan):
                errors.append(f"{prefix}: Invalid PAN format '{e.pan}'")

            # Section validation
            if e.section and e.section.upper() not in _KNOWN_SECTIONS:
                warnings.append(f"{prefix}: Unknown TDS section '{e.section}'")

            # TDS amount sanity
            if e.amount_paid > 0 and e.tds_deducted == 0:
                warnings.append(
                    f"{prefix}: Amount paid ₹{e.amount_paid:,.2f} but TDS = 0"
                )

            # TDS > payment
            if e.tds_deducted > e.amount_paid:
                errors.append(
                    f"{prefix}: TDS ({e.tds_deducted:,.2f}) > Amount paid ({e.amount_paid:,.2f})"
                )

            # Rate check
            if e.rate > 0 and e.amount_paid > 0:
                expected = round(e.amount_paid * e.rate / 100, 2)
                e.expected_tds = expected
                diff = abs(e.tds_deducted - expected)
                if diff > 1.0:
                    warnings.append(
                        f"{prefix}: TDS mismatch — expected ₹{expected:,.2f} "
                        f"at {e.rate}%, got ₹{e.tds_deducted:,.2f}"
                    )
                e.tds_shortfall = max(0.0, expected - e.tds_deducted)

            # Deposited vs deducted
            if e.tds_deposited > 0 and e.tds_deposited < e.tds_deducted:
                errors.append(
                    f"{prefix}: TDS deposited ({e.tds_deposited:,.2f}) < "
                    f"deducted ({e.tds_deducted:,.2f})"
                )

        return len(errors) == 0, errors, warnings
