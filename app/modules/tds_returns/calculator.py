"""
TDS Returns Calculator.

Assigns Challan Serial Numbers and computes group summaries.

Challan Assignment Rules (from desired output analysis):
  - Each unique (source_file, source_group, payment_date, section) = one Challan
  - Challans are numbered sequentially sorted by:
      1. section code  (alphabetical: 194A → 194C → 194J → 194Q …)
      2. payment_date  (chronological: Oct → Nov → Dec)
      3. source_group  (stable tie-break for same section+date sub-groups)
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


def _date_sort_key(date_str: str) -> tuple:
    """
    Convert DD/MM/YYYY → (YYYY, MM, DD) tuple for chronological sorting.
    Returns (0, 0, 0) for empty / unparseable strings so they sort first.
    """
    if not date_str:
        return (0, 0, 0)
    parts = date_str.split('/')
    if len(parts) == 3:
        try:
            return (int(parts[2]), int(parts[1]), int(parts[0]))
        except ValueError:
            pass
    return (0, 0, 0)


class TDSReturnsCalculator:

    def assign_challans(self, entries: List[TDSEntry]) -> List[ChallanGroup]:
        """
        Assign challan serial numbers to entries.

        Groups by (source_file, source_group, payment_date, section).
        Each unique combination = one challan.

        Serial numbers are assigned after sorting groups by:
          (section ASC, payment_date ASC, source_group ASC)

        This ensures 194A challans always precede 194C, which precede 194Q,
        and within each section the months run Oct → Nov → Dec — matching
        the expected Form 26Q output ordering.

        Entry order within each challan group is preserved (file-appearance order)
        so individual deductee rows stay in their original sequence.
        """
        # Pass 1 — collect unique challan keys and accumulate their entries,
        # preserving file-appearance order within each group.
        key_entries: Dict[Tuple, List[TDSEntry]] = {}
        key_order: List[Tuple] = []   # deduplicated insertion-order list

        for entry in entries:
            key = (entry.source_file, entry.source_group,
                   entry.payment_date_str, entry.section)
            if key not in key_entries:
                key_entries[key] = []
                key_order.append(key)
            key_entries[key].append(entry)

        # Pass 2 — sort the unique challan keys.
        #   Primary   : section code (alphabetical)
        #   Secondary : payment date (chronological DD/MM/YYYY)
        #   Tertiary  : source_group (stable tie-break for same section+date sub-groups,
        #               e.g.  SDM-11.xls_194C  before  SDM-11.xls_194C-2)
        def _sort_key(k: Tuple) -> Tuple:
            _src_file, src_group, pay_date, section = k
            return (section, _date_sort_key(pay_date), src_group)

        sorted_keys = sorted(key_order, key=_sort_key)

        # Pass 3 — assign serial numbers in sorted order and build ChallanGroup objects.
        groups: List[ChallanGroup] = []
        for serial, key in enumerate(sorted_keys, start=1):
            _src_file, _src_group, pay_date, section = key
            grp = ChallanGroup(
                challan_serial=serial,
                section=section,
                payment_date=pay_date,
            )
            for entry in key_entries[key]:
                entry.challan_serial = serial
                grp.entries.append(entry)
                grp.total_amount += entry.amount_paid
                grp.total_tds += entry.tds_deducted
                grp.deductee_count += 1
            groups.append(grp)

        logger.info(
            "Assigned %d challans across %d entries",
            len(groups), len(entries)
        )
        return groups

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
