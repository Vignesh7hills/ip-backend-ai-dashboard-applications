"""Trial Balance Service — parse → validate → calculate → generate Excel.

Supports SINGLE file OR MULTIPLE files merged into ONE output.
"""
import time
from typing import List
from app.modules.trial_balance.parser import TrialBalanceParser, TrialBalanceEntry
from app.modules.trial_balance.validator import TrialBalanceValidator
from app.modules.trial_balance.calculator import TrialBalanceCalculator
from app.modules.trial_balance.generator import TrialBalanceExcelGenerator
from app.core.logger import get_logger
from app.core.exceptions import EmptyFileError

logger = get_logger(__name__)


class TrialBalanceService:

    def __init__(self):
        self.parser     = TrialBalanceParser()
        self.validator  = TrialBalanceValidator()
        self.calculator = TrialBalanceCalculator()
        self.generator  = TrialBalanceExcelGenerator()

    def process(self, file_path: str, company_name: str = '') -> dict:
        """Single-file wrapper — kept for backward compatibility."""
        return self.process_multiple([file_path], company_name=company_name)

    def process_multiple(
        self,
        file_paths: List[str],
        company_name: str = '',
    ) -> dict:
        """
        Parse one OR many BS/PL/TB files and produce a SINGLE merged output.

        Multiple files are parsed independently; their entries are combined,
        validated together, and rendered once — so uploading a Balance Sheet
        and a P&L together collapses into one balanced Trial Balance.
        """
        t0 = time.perf_counter()
        logs: list  = []
        warnings: list = []
        errors: list   = []

        def log(msg):
            logger.info(msg)
            logs.append({'level': 'INFO', 'message': msg})

        if isinstance(file_paths, str):
            file_paths = [file_paths]

        log(f"Step 1/4: Parsing {len(file_paths)} file(s)")
        all_entries: List[TrialBalanceEntry] = []

        for fp in file_paths:
            fname = fp.rsplit('/', 1)[-1]
            try:
                entries = self.parser.parse_file(fp)
                all_entries.extend(entries)
                log(f"  {fname}: {len(entries)} entries")
            except Exception as exc:
                warnings.append(f"[{fname}] could not be parsed: {exc}")
                logger.warning("Failed to parse %s: %s", fname, exc)

        log(f"Parsed {len(all_entries)} total entries from {len(file_paths)} file(s)")

        if not all_entries:
            raise EmptyFileError("No data found in any of the supplied files.")

        # ── Closing Stock exclusion ───────────────────────────────────────────
        # Per requirement: if Closing Stock already appears in a Balance Sheet
        # or Trading Account input, remove it from the combined TB to avoid
        # double-counting.  We detect this by checking whether any entry is
        # tagged as CLOSING STOCK group.
        has_closing = any(
            'closing' in (e.group or '').lower() or
            'closing' in e.account_name.lower()
            for e in all_entries
        )
        if has_closing:
            before = len(all_entries)
            all_entries = [
                e for e in all_entries
                if not ('closing' in (e.group or '').lower() or
                        'closing' in e.account_name.lower())
            ]
            log(f"Excluded {before - len(all_entries)} Closing Stock entries (per TB rules)")
            warnings.append(
                "Closing Stock excluded from Trial Balance — "
                "already captured in Balance Sheet / Trading Account."
            )

        # ── Net Profit → Capital transfer ─────────────────────────────────────
        # Per requirement: if Dr ≠ Cr after merging all files, the difference
        # represents Net Profit / Loss for the year.  Transfer it to Capital A/c
        # as a Debit entry so the TB balances.
        total_dr = sum(e.debit  for e in all_entries)
        total_cr = sum(e.credit for e in all_entries)
        np_diff  = total_cr - total_dr          # positive = profit, negative = loss

        if abs(np_diff) > 0.50:
            from app.modules.trial_balance.parser import TrialBalanceEntry as TBE
            np_entry = TBE(
                account_name='Net Profit for the Year',
                group='CAPITAL',
            )
            if np_diff > 0:
                # Profit → debit Capital (closes P&L credit surplus)
                np_entry.debit  = round(np_diff, 2)
                np_entry.credit = 0.0
            else:
                # Loss → credit Capital
                np_entry.credit = round(abs(np_diff), 2)
                np_entry.debit  = 0.0
            all_entries.append(np_entry)
            log(
                f"Net {'Profit' if np_diff > 0 else 'Loss'} of "
                f"{abs(np_diff):,.2f} transferred to CAPITAL A/c"
            )
            warnings.append(
                f"Net {'Profit' if np_diff > 0 else 'Loss'} ₹{abs(np_diff):,.2f} "
                f"transferred to Capital A/c per TB rules."
            )

        log("Step 2/4: Validating")
        is_valid, v_errors, v_warnings = self.validator.validate(all_entries)
        errors.extend(v_errors)
        warnings.extend(v_warnings)

        log("Step 3/4: Calculating group totals")
        groups = self.calculator.compute(all_entries)
        log(f"Computed {len(groups)} groups")

        if not groups:
            raise EmptyFileError("No data groups found in trial balance.")

        log("Step 4/4: Generating Excel")
        excel_bytes = self.generator.generate(
            groups,
            entries_flat=all_entries,
            company_name=company_name,
            validation_errors=errors,
            warnings=warnings,
        )

        duration_ms = (time.perf_counter() - t0) * 1000
        log(f"Completed in {duration_ms:.0f} ms")

        return {
            'excel_bytes':  excel_bytes,
            'records':      sum(len(g.entries) for g in groups.values()),
            'warnings':     warnings,
            'errors':       errors,
            'logs':         logs,
            'duration_ms':  duration_ms,
        }
