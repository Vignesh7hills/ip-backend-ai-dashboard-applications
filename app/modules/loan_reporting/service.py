"""
Loan Reporting Service — orchestrates the full pipeline:
  parse → normalize → validate → calculate → generate Excel
"""
import time
from typing import Optional
from app.modules.loan_reporting.parser import LedgerPDFParser
from app.modules.loan_reporting.normalizer import LoanReportingNormalizer
from app.modules.loan_reporting.validator import LoanReportingValidator
from app.modules.loan_reporting.calculator import LoanReportingCalculator
from app.modules.loan_reporting.generator import LoanReportingExcelGenerator
from app.core.logger import get_logger
from app.core.exceptions import EmptyFileError, FileParseError

logger = get_logger(__name__)


class LoanReportingService:

    def __init__(self):
        self.parser = LedgerPDFParser()
        self.normalizer = LoanReportingNormalizer()
        self.validator = LoanReportingValidator()
        self.calculator = LoanReportingCalculator()
        self.generator = LoanReportingExcelGenerator()

    def process(
        self,
        pdf_path: str,
        company_name: str = '',
    ) -> dict:
        """
        Full pipeline. Returns:
        {
            "excel_bytes": bytes,
            "records": int,
            "warnings": [...],
            "errors": [...],
            "logs": [...],
            "duration_ms": float,
        }
        """
        t0 = time.perf_counter()
        logs = []
        warnings = []
        errors = []

        def log(msg: str, level: str = 'INFO'):
            logger.info(msg) if level == 'INFO' else logger.warning(msg)
            logs.append({'level': level, 'message': msg})

        # ── Step 1: Parse ─────────────────────────────────────────────────────
        log("Step 1/5: Parsing PDF ledger")
        raw_accounts = self.parser.parse_file(pdf_path)
        if not raw_accounts:
            raise EmptyFileError("No account blocks found in the PDF.")
        log(f"Parsed {len(raw_accounts)} account blocks")

        # ── Step 2: Normalize ─────────────────────────────────────────────────
        log("Step 2/5: Normalizing transactions")
        norm_accounts = self.normalizer.normalize(raw_accounts)
        log(f"Normalized {len(norm_accounts)} accounts")

        # ── Step 3: Validate ──────────────────────────────────────────────────
        log("Step 3/5: Validating data")
        validated_accounts = self.validator.validate_all(norm_accounts)
        for acct in validated_accounts:
            if acct.validation_errors:
                for err in acct.validation_errors:
                    errors.append(f"[{acct.name}] {err}")
            if not acct.balance_validated:
                warnings.append(f"[{acct.name}] Balance validation failed — results may be inaccurate")

        # ── Step 4: Calculate ─────────────────────────────────────────────────
        log("Step 4/5: Calculating report metrics")
        report_rows = self.calculator.compute_report_rows(validated_accounts)
        log(f"Report has {len(report_rows)} qualifying accounts")

        if not report_rows:
            raise EmptyFileError(
                "No qualifying loan accounts found. "
                "Ensure the file contains Loans and Advances accounts with credit balances."
            )

        # ── Step 5: Generate Excel ────────────────────────────────────────────
        log("Step 5/5: Generating Excel report")
        excel_bytes = self.generator.generate(report_rows, company_name=company_name)

        duration_ms = (time.perf_counter() - t0) * 1000
        log(f"Pipeline completed in {duration_ms:.0f} ms")

        return {
            'excel_bytes': excel_bytes,
            'records': len(report_rows),
            'warnings': warnings,
            'errors': errors,
            'logs': logs,
            'duration_ms': duration_ms,
        }
