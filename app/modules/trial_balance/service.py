"""Trial Balance Service — AUTO-ADAPTIVE pipeline. Accepts any file format/content."""
import time
from typing import List
from app.utils.universal_parser import parse_for_trial_balance
from app.modules.trial_balance.parser import TrialBalanceEntry
from app.modules.trial_balance.validator import TrialBalanceValidator
from app.modules.trial_balance.calculator import TrialBalanceCalculator, compute_pl_net_profit
from app.modules.trial_balance.generator import TrialBalanceExcelGenerator
from app.core.logger import get_logger
from app.core.exceptions import EmptyFileError

logger = get_logger(__name__)


class TrialBalanceService:

    def __init__(self):
        self.validator  = TrialBalanceValidator()
        self.calculator = TrialBalanceCalculator()
        self.generator  = TrialBalanceExcelGenerator()

    def process(self, file_path: str, company_name: str = '') -> dict:
        return self.process_multiple([file_path], company_name=company_name)

    def process_multiple(self, file_paths: List[str], company_name: str = '') -> dict:
        t0 = time.perf_counter()
        logs, warnings, errors = [], [], []

        def log(msg):
            logger.info(msg)
            logs.append({'level': 'INFO', 'message': msg})

        if isinstance(file_paths, str):
            file_paths = [file_paths]

        log(f"Step 1/4: Parsing {len(file_paths)} file(s) — auto-adaptive")
        all_entries: List[TrialBalanceEntry] = []

        for fp in file_paths:
            fname = fp.rsplit('/', 1)[-1]
            try:
                entries, parse_warnings = parse_for_trial_balance(fp)
                warnings.extend(parse_warnings)
                all_entries.extend(entries)
                log(f"  {fname}: {len(entries)} entries")
            except Exception as exc:
                warnings.append(f"[{fname}] could not be parsed: {exc}")
                logger.warning("Failed: %s: %s", fname, exc)

        log(f"Total: {len(all_entries)} entries")

        if not all_entries:
            raise EmptyFileError("No data found in any of the supplied files.")

        # ── Closing Stock exclusion ───────────────────────────────────────────
        has_closing = any(
            'closing' in (e.group or '').lower() or 'closing' in e.account_name.lower()
            for e in all_entries
        )
        if has_closing:
            before = len(all_entries)
            all_entries = [e for e in all_entries
                           if not ('closing' in (e.group or '').lower()
                                   or 'closing' in e.account_name.lower())]
            log(f"Excluded {before - len(all_entries)} Closing Stock entries")
            warnings.append("Closing Stock excluded from Trial Balance (per TB rules).")

        # ── Trial balance integrity check ─────────────────────────────────────
        # A correctly-entered trial balance MUST have equal Debit and Credit
        # totals. Any leftover Dr/Cr difference is a DATA ERROR in the source —
        # it is NOT profit and must never be silently plugged into Capital.
        total_dr = round(sum(e.debit  for e in all_entries), 2)
        total_cr = round(sum(e.credit for e in all_entries), 2)
        imbalance = round(total_dr - total_cr, 2)

        if abs(imbalance) > 0.50:
            msg = (
                f"Trial balance does NOT tie: Debit \u20b9{total_dr:,.2f} vs "
                f"Credit \u20b9{total_cr:,.2f} (difference \u20b9{abs(imbalance):,.2f}). "
                f"This is an error in the source data, not a profit \u2014 please "
                f"correct it. No balancing entry has been created."
            )
            warnings.append(msg)
            log("\u26a0 " + msg)

        # ── Net Profit / (Loss) — derived from the P&L GROUPS, never the Dr/Cr gap ──
        pl_groups = self.calculator.compute(all_entries)
        net_pl = compute_pl_net_profit(pl_groups)
        if abs(net_pl) > 0.50:
            tag = 'Profit' if net_pl > 0 else 'Loss'
            log(f"Net {tag} \u20b9{abs(net_pl):,.2f} (computed from P&L ledgers)")
            warnings.append(
                f"Net {tag} \u20b9{abs(net_pl):,.2f} computed from P&L ledgers "
                f"(shown in the Summary sheet)."
            )

        log("Step 2/4: Validating")
        _, v_errors, v_warnings = self.validator.validate(all_entries)
        errors.extend(v_errors)
        warnings.extend(v_warnings)

        log("Step 3/4: Calculating groups")
        groups = self.calculator.compute(all_entries)
        log(f"Computed {len(groups)} groups")

        if not groups:
            raise EmptyFileError("No data groups found in trial balance.")

        log("Step 4/4: Generating Excel")
        excel_bytes = self.generator.generate(
            groups, entries_flat=all_entries,
            company_name=company_name,
            validation_errors=errors, warnings=warnings,
        )

        duration_ms = (time.perf_counter() - t0) * 1000
        log(f"Completed in {duration_ms:.0f} ms")

        return {
            'excel_bytes': excel_bytes,
            'records':     sum(len(g.entries) for g in groups.values()),
            'warnings':    warnings,
            'errors':      errors,
            'logs':        logs,
            'duration_ms': duration_ms,
        }
