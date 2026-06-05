"""
Loan Reporting Service — AUTO-ADAPTIVE pipeline.

Accepts ANY file format AND ANY content type:
  - Tally PDF/Excel ledger (native)
  - TDS register → extracts party names, PANs, amounts
  - Trial Balance → extracts account balances
  - CSV/TXT with any column layout

NEVER returns 422. Always produces output with warnings for inferred/missing data.
"""
import os
import time
from typing import List
from app.utils.universal_parser import parse_for_loan_reporting
from app.modules.loan_reporting.normalizer import LoanReportingNormalizer
from app.modules.loan_reporting.validator import LoanReportingValidator
from app.modules.loan_reporting.calculator import LoanReportingCalculator
from app.modules.loan_reporting.generator import LoanReportingExcelGenerator
from app.core.logger import get_logger
from app.core.exceptions import EmptyFileError, FileParseError

logger = get_logger(__name__)


class LoanReportingService:

    def __init__(self):
        self.normalizer  = LoanReportingNormalizer()
        self.validator   = LoanReportingValidator()
        self.calculator  = LoanReportingCalculator()
        self.generator   = LoanReportingExcelGenerator()

    def process(self, pdf_path: str, company_name: str = '',
                file_type: str = '') -> dict:
        t0 = time.perf_counter()
        logs, warnings, errors = [], [], []

        def log(msg, level='INFO'):
            logger.info(msg) if level == 'INFO' else logger.warning(msg)
            logs.append({'level': level, 'message': msg})

        # ── Step 1: Parse (auto-adaptive) ─────────────────────────────────────
        log("Step 1/5: Auto-detecting and parsing file")
        raw_accounts, parse_warnings = parse_for_loan_reporting(pdf_path)
        warnings.extend(parse_warnings)

        if not raw_accounts:
            raise EmptyFileError("No account data found in file.")

        log(f"Extracted {len(raw_accounts)} account(s) from file")
        for w in parse_warnings:
            log(f"  [note] {w}", level='WARN')

        # ── Step 2: Normalize ─────────────────────────────────────────────────
        log("Step 2/5: Normalizing transactions")
        norm_accounts = self.normalizer.normalize(raw_accounts)
        log(f"Normalized {len(norm_accounts)} accounts")

        # ── Step 3: Validate ──────────────────────────────────────────────────
        log("Step 3/5: Validating")
        validated = self.validator.validate_all(norm_accounts)
        for acct in validated:
            for err in acct.validation_errors:
                errors.append(f"[{acct.name}] {err}")
            if not acct.balance_validated:
                warnings.append(f"[{acct.name}] Balance validation skipped (limited data)")

        # ── Step 4: Calculate ─────────────────────────────────────────────────
        log("Step 4/5: Calculating report metrics")
        report_rows = self.calculator.compute_report_rows(validated)
        log(f"Report: {len(report_rows)} qualifying accounts")

        # If calculator excluded all accounts (all Dr-only), relax the filter
        if not report_rows:
            log("No Cr-balance accounts found — including all accounts", level='WARN')
            warnings.append(
                "No credit-balance accounts detected. All accounts shown "
                "(typical when file contains payment records only)."
            )
            report_rows = self._force_report_rows(validated)

        if not report_rows:
            raise EmptyFileError(
                "No reportable accounts found in this file."
            )

        # ── Step 5: Generate Excel ────────────────────────────────────────────
        log("Step 5/5: Generating Excel report")
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
        """Fallback: create report rows directly from account data."""
        from app.models.account import LoanReportRow
        from app.utils.amount_parser import cr_value

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
