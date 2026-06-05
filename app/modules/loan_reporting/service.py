"""
Loan Reporting Service — multi-file, annexure-enriched pipeline.

Accepts:
  - Multiple ledger files (any format) — merged into one report
  - Optional annexure_ref dict from parse_annexure_files — PAN/Address enrichment
"""
import os
import time
from typing import List, Dict, Optional
from app.utils.universal_parser import parse_for_loan_reporting
from app.utils.annexure_parser import enrich_with_annexure
from app.modules.loan_reporting.normalizer import LoanReportingNormalizer
from app.modules.loan_reporting.validator import LoanReportingValidator
from app.modules.loan_reporting.calculator import LoanReportingCalculator
from app.modules.loan_reporting.generator import LoanReportingExcelGenerator
from app.core.logger import get_logger
from app.core.exceptions import EmptyFileError

logger = get_logger(__name__)


class LoanReportingService:

    def __init__(self):
        self.normalizer  = LoanReportingNormalizer()
        self.validator   = LoanReportingValidator()
        self.calculator  = LoanReportingCalculator()
        self.generator   = LoanReportingExcelGenerator()

    def process(self, pdf_path: str, company_name: str = '',
                file_type: str = '', annexure_ref: Dict = None) -> dict:
        """Single-file convenience wrapper."""
        return self.process_multiple(
            file_paths=[pdf_path],
            company_name=company_name,
            annexure_ref=annexure_ref or {},
        )

    def process_multiple(self, file_paths: List[str], company_name: str = '',
                         annexure_ref: Dict = None) -> dict:
        """
        Full pipeline for one or more ledger files.
        All accounts across all files are merged into a single report.
        """
        t0 = time.perf_counter()
        logs, warnings, errors = [], [], []
        annexure_ref = annexure_ref or {}

        def log(msg, level='INFO'):
            logger.info(msg) if level == 'INFO' else logger.warning(msg)
            logs.append({'level': level, 'message': msg})

        if isinstance(file_paths, str):
            file_paths = [file_paths]

        # ── Step 1: Parse all files ───────────────────────────────────────────
        log(f"Step 1/5: Parsing {len(file_paths)} file(s)")
        all_raw_accounts = []

        for fp in file_paths:
            fname = fp.rsplit('/', 1)[-1]
            try:
                raw_accounts, parse_warnings = parse_for_loan_reporting(fp)
                warnings.extend(parse_warnings)
                all_raw_accounts.extend(raw_accounts)
                log(f"  {fname}: {len(raw_accounts)} account(s)")
                for w in parse_warnings:
                    log(f"    [note] {w}", level='WARN')
            except Exception as exc:
                warnings.append(f"[{fname}] parse failed: {exc}")
                log(f"  {fname}: FAILED — {exc}", level='WARN')

        if not all_raw_accounts:
            raise EmptyFileError("No account data found in any of the supplied files.")

        log(f"Total: {len(all_raw_accounts)} account(s) from {len(file_paths)} file(s)")

        # ── Step 2: Normalize ─────────────────────────────────────────────────
        log("Step 2/5: Normalizing")
        norm_accounts = self.normalizer.normalize(all_raw_accounts)

        # ── Step 3: Validate ──────────────────────────────────────────────────
        log("Step 3/5: Validating")
        validated = self.validator.validate_all(norm_accounts)
        for acct in validated:
            for err in acct.validation_errors:
                errors.append(f"[{acct.name}] {err}")
            if not acct.balance_validated:
                warnings.append(f"[{acct.name}] Balance check skipped")

        # ── Step 4: Calculate ─────────────────────────────────────────────────
        log("Step 4/5: Calculating")
        report_rows = self.calculator.compute_report_rows(validated)
        log(f"Report: {len(report_rows)} qualifying accounts")

        if not report_rows:
            warnings.append("No Cr-balance accounts — showing all accounts.")
            report_rows = self._force_report_rows(validated)

        if not report_rows:
            raise EmptyFileError("No reportable accounts found.")

        # ── Step 5: Enrich with Annexure PAN/Address ──────────────────────────
        if annexure_ref:
            log(f"Step 5a: Enriching {len(report_rows)} rows from {len(annexure_ref)} annexure records")
            report_rows = enrich_with_annexure(report_rows, annexure_ref)
            enriched = sum(1 for r in report_rows if getattr(r, 'pan', '') or getattr(r, 'address', ''))
            log(f"  Enriched: {enriched}/{len(report_rows)} rows have PAN/Address")

        # ── Step 6: Generate Excel ────────────────────────────────────────────
        log("Step 5/5: Generating Excel")
        excel_bytes = self.generator.generate(report_rows, company_name=company_name)

        duration_ms = (time.perf_counter() - t0) * 1000
        log(f"Completed in {duration_ms:.0f} ms")

        return {
            'excel_bytes': excel_bytes,
            'records':     len(report_rows),
            'warnings':    warnings,
            'errors':      errors,
            'logs':        logs,
            'duration_ms': duration_ms,
        }

    def _force_report_rows(self, validated_accounts) -> list:
        from app.models.account import LoanReportRow
        rows = []
        for i, acct in enumerate(validated_accounts, 1):
            taken  = sum(t.credit for t in acct.transactions if not t.is_opening)
            repaid = sum(t.debit  for t in acct.transactions if not t.is_opening)
            if taken == 0 and repaid == 0:
                taken  = acct.opening_balance if acct.opening_direction == 'Cr' else 0
                repaid = acct.opening_balance if acct.opening_direction == 'Dr' else 0
            maximum = max(taken, acct.closing_balance, 0.0)
            rows.append(LoanReportRow(
                sr_no=i,
                name=acct.name,
                address=acct.location or '',
                pan=acct.pan or '',
                taken=taken,
                repaid=repaid,
                maximum=maximum,
                squared_up='YES' if acct.closing_balance == 0 else 'NO',
            ))
        return rows
