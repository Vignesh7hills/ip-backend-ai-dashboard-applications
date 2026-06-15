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
                f_dr = round(sum(e.debit for e in entries), 2)
                f_cr = round(sum(e.credit for e in entries), 2)
                log(f"  {fname}: {len(entries)} entries "
                    f"(Dr \u20b9{f_dr:,.2f} / Cr \u20b9{f_cr:,.2f})")
                # A published BS or P&L always balances internally; a per-file
                # gap means extraction lost or double-counted rows — UNLESS the
                # gap is the year's net profit (P&L files commonly omit it as a
                # ledger line; the pipeline transfers it to CAPITAL later).
                f_diff = round(f_dr - f_cr, 2)
                if abs(f_diff) > 1.0:
                    warnings.append(
                        f"[{fname}] parsed Dr \u20b9{f_dr:,.2f} vs Cr \u20b9{f_cr:,.2f} "
                        f"(difference \u20b9{abs(f_diff):,.2f}). If this is not the "
                        f"net profit/loss for the year, please verify the file."
                    )
            except Exception as exc:
                warnings.append(f"[{fname}] could not be parsed: {exc}")
                logger.warning("Failed: %s: %s", fname, exc)

        log(f"Total: {len(all_entries)} entries")

        if not all_entries:
            raise EmptyFileError("No data found in any of the supplied files.")

        # ── Net Profit / (Loss) — computed BEFORE Closing Stock exclusion ────
        # Closing Stock must be excluded from the TB OUTPUT sheet (per TB
        # rules), but it is part of the P&L equation. Computing net profit
        # after the exclusion silently forced closing = 0 and produced a
        # fictitious loss. So: compute P&L first on the FULL entry list.
        pl_groups = self.calculator.compute(all_entries)
        net_pl = compute_pl_net_profit(pl_groups)
        if abs(net_pl) > 0.50:
            tag = 'Profit' if net_pl > 0 else 'Loss'
            log(f"Net {tag} \u20b9{abs(net_pl):,.2f} (computed from P&L ledgers, incl. stock)")
            warnings.append(
                f"Net {tag} \u20b9{abs(net_pl):,.2f} computed from P&L ledgers "
                f"(shown in the Summary sheet)."
            )

        # ── Closing Stock exclusion (OUTPUT sheet only) ───────────────────────
        # The client rule (Note_TB point 2): do NOT include Closing Stock in
        # the TB output when it appears in the Balance Sheet or Trading Account.
        #
        # We filter at TWO levels:
        # 1. Entry-level: remove entries whose name/group directly says "closing"
        #    (catches the most common case: explicit Closing Stock entries).
        # 2. After-calculator: the calculator maps some stock groups (e.g.
        #    "Cloth Stock", "Yarn Stock" from a Format-C collapse) to
        #    'CLOSING STOCK'. We run a second pass after computing groups to
        #    strip any entry that ends up in the CLOSING STOCK group.
        #
        # Known stock account sub-names that are always Closing Stock (never
        # Opening Stock) when seen on the Credit/Balance Sheet side:
        _CLOSING_STOCK_NAMES = {
            'cloth stock', 'yarn stock', 'stock in hand', 'stock-in-hand',
            'closing stock', 'closing stock a/c', 'closing stock ac',
        }

        def _is_closing_stock_entry(e: TrialBalanceEntry) -> bool:
            nm  = e.account_name.lower().strip()
            grp = (e.group or '').lower().strip()
            # Explicit "closing" keyword in name or group
            if 'closing' in nm or 'closing' in grp:
                return True
            # Named stock-in-hand accounts that are always Closing Stock in a BS/TB.
            # In the BS these appear as Dr (Current Assets), in the P&L as Cr.
            # Either way they must be excluded per the client rule.
            # We also check 'cloth stock' / 'stock in hand' type groups that the
            # calculator maps to CLOSING STOCK internally.
            _STOCK_GROUPS = {
                'cloth stock', 'yarn stock', 'stock in hand', 'stock-in-hand',
                'closing stock', 'closing stock a/c', 'closing stock ac',
                'stock', 'stock a/c',
            }
            if nm in _STOCK_GROUPS or grp in _STOCK_GROUPS:
                # Only exclude if it's NOT labelled as Opening Stock
                if 'opening' not in nm and 'opening' not in grp:
                    return True
            return False

        has_closing = any(_is_closing_stock_entry(e) for e in all_entries)
        if has_closing:
            before = len(all_entries)
            all_entries = [e for e in all_entries if not _is_closing_stock_entry(e)]
            log(f"Excluded {before - len(all_entries)} Closing Stock entries")
            warnings.append("Closing Stock excluded from Trial Balance (per TB rules).")

        # ── Trial balance integrity check ─────────────────────────────────────
        total_dr = round(sum(e.debit  for e in all_entries), 2)
        total_cr = round(sum(e.credit for e in all_entries), 2)
        imbalance = round(total_dr - total_cr, 2)

        # ── Net Profit → Capital plug (Note_TB rule 6) ────────────────────────
        # "If difference is only to the tune of net profit amount, transfer
        #  net profit to capital a/c in trial balance with debit column."
        # Capital already includes the year's profit, so the TB is short on
        # the Dr side by exactly the net profit. Plug ONLY when the gap
        # matches the independently computed P&L result (tolerance \u20b91).
        if abs(imbalance) > 0.50 and abs(net_pl) > 0.50 \
                and abs(abs(imbalance) - abs(net_pl)) <= 1.00:
            plug = TrialBalanceEntry(account_name='NET PROFIT', group='CAPITAL')
            if imbalance < 0:          # Cr > Dr -> plug on the Debit side
                plug.debit = abs(imbalance)
            else:                      # Dr > Cr -> plug on the Credit side
                plug.account_name = 'NET LOSS'
                plug.credit = abs(imbalance)
            all_entries.append(plug)
            log(f"\u2713 {plug.account_name} \u20b9{abs(imbalance):,.2f} "
                f"transferred to CAPITAL (matches P&L result) \u2014 TB now tallies")
            warnings.append(
                f"{plug.account_name} \u20b9{abs(imbalance):,.2f} transferred to "
                f"CAPITAL A/c per TB rules (Dr/Cr gap equals computed P&L result)."
            )
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