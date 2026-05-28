from dataclasses import dataclass, field
from datetime import date
from typing import Literal, Optional


@dataclass
class RawTransaction:
    """
    A single parsed transaction row from the ledger PDF,
    before normalization/enrichment.
    """
    date_str: str
    voucher_type: str            # Opbl | Pay | Rec | Jrn
    voucher_no: str              # BK-189, 885, etc.
    description: str
    narration: str
    amount: float                # raw amount from the PDF column
    balance_str: str             # "9,41,800.00Cr" | "7,25,007.00Dr" | "-"
    balance: float               # numeric value
    balance_direction: str       # "Cr" | "Dr" | "Zero"


@dataclass
class NormalizedTransaction:
    """
    A normalized transaction with resolved debit/credit columns,
    typed date, and classified voucher kind.
    """
    txn_date: Optional[date]
    voucher_type: str            # Opbl | Pay | Rec | Jrn
    voucher_no: str
    description: str
    narration: str
    debit: float                 # 0.0 if not a debit transaction
    credit: float                # 0.0 if not a credit transaction
    balance: float
    balance_direction: str       # Cr | Dr | Zero
    is_bank_txn: bool            # True for BK-prefix vouchers
    is_journal: bool             # True for Jrn vouchers
    is_opening: bool             # True for Opbl
    journal_no: str              # e.g. "885" for grouping same-journal entries
