from dataclasses import dataclass, field
from typing import Optional
from app.models.transaction import NormalizedTransaction


@dataclass
class NormalizedAccount:
    """
    Fully parsed and normalized ledger account ready for calculations.
    """
    name: str
    location: str
    pan: str                     # populated from external source if available
    opening_balance: float
    opening_direction: str       # Cr | Dr | Zero
    closing_balance: float
    closing_direction: str       # Cr | Dr | Zero
    total_debit: float
    total_credit: float
    transactions: list = field(default_factory=list)  # list[NormalizedTransaction]

    # Populated by validator
    balance_validated: bool = False
    validation_errors: list = field(default_factory=list)


@dataclass
class LoanReportRow:
    """
    Single row of the Loan Reporting output (Form 3CD Clause 31).
    """
    sr_no: int
    name: str
    address: str
    pan: str
    taken: float         # total Rec (bank) credits during year
    repaid: float        # total Pay (bank) debits during year
    maximum: float       # max Cr balance at any real checkpoint
    squared_up: str      # YES | NO
