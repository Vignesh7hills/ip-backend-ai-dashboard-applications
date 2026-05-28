"""
TDS Returns Calculator.

Assigns Challan Serial Numbers and computes group summaries.

Challan Assignment Rules (from desired output analysis):
  - Each unique (source_group, payment_date, section) combination = one Challan
  - Challans are numbered sequentially in the order they appear across all files
  - Within a single file, each TDS group section = one challan
  - For Vanshika: each file + section = one challan (94c oct, 94c nov, 94q dec, etc.)
  - For Shridhar: each TDS group within a file (TYPE 01 group, TYPE 02 group) = one challan
"""
from typing import List, Dict, Tuple
from dataclasses import dataclass, field
from app.modules.tds_returns.parser import TDSEntry
from app.core.logger import get_logger

logger = get_logger(__name__)


@dataclass
class ChallanGroup:
    challan_serial: int
    section: str
    payment_date: str
    entries: List[TDSEntry] = field(default_factory=list)
    total_amount: float = 0.0
    total_tds: float = 0.0
    deductee_count: int = 0


@dataclass
class SectionSummary:
    """For backward compatibility with service/validator."""
    section: str
    entries: List[TDSEntry] = field(default_factory=list)
    total_amount_paid: float = 0.0
    total_tds_deducted: float = 0.0
    total_tds_deposited: float = 0.0
    total_shortfall: float = 0.0
    deductee_count: int = 0


class TDSReturnsCalculator:

    def assign_challans(self, entries: List[TDSEntry]) -> List[ChallanGroup]:
        """
        Assign challan serial numbers to entries.
        Groups by (source_file, source_group, payment_date, section).
        Each unique combination gets a sequential challan number.
        """
        # Build ordered list of unique challan keys
        seen_keys: Dict[Tuple, int] = {}
        challan_counter = 0
        groups: Dict[int, ChallanGroup] = {}

        for entry in entries:
            key = (entry.source_file, entry.source_group, entry.payment_date_str, entry.section)
            if key not in seen_keys:
                challan_counter += 1
                seen_keys[key] = challan_counter
                groups[challan_counter] = ChallanGroup(
                    challan_serial=challan_counter,
                    section=entry.section,
                    payment_date=entry.payment_date_str,
                )

            challan_no = seen_keys[key]
            entry.challan_serial = challan_no
            grp = groups[challan_no]
            grp.entries.append(entry)
            grp.total_amount += entry.amount_paid
            grp.total_tds += entry.tds_deducted
            grp.deductee_count += 1

        logger.info(
            "Assigned %d challans across %d entries",
            len(groups), len(entries)
        )
        return [groups[i] for i in sorted(groups.keys())]

    def compute(self, entries: List[TDSEntry]) -> Dict[str, SectionSummary]:
        """Section-wise summary (for backward compatibility)."""
        sections: Dict[str, SectionSummary] = {}
        for entry in entries:
            sec = entry.section.upper() if entry.section else 'UNKNOWN'
            if sec not in sections:
                sections[sec] = SectionSummary(section=sec)
            s = sections[sec]
            s.entries.append(entry)
            s.total_amount_paid += entry.amount_paid
            s.total_tds_deducted += entry.tds_deducted
            s.total_tds_deposited += entry.tds_deposited
            s.total_shortfall += entry.tds_shortfall
            s.deductee_count += 1
        return sections
