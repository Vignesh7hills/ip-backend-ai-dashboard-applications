"""TDS Returns Service — AUTO-ADAPTIVE pipeline. Accepts any file format/content."""
import time
from typing import List
from app.utils.universal_parser import parse_for_tds_returns
from app.modules.tds_returns.validator import TDSReturnsValidator
from app.modules.tds_returns.calculator import TDSReturnsCalculator
from app.modules.tds_returns.generator import TDSReturnsExcelGenerator
from app.core.logger import get_logger
from app.core.exceptions import EmptyFileError

logger = get_logger(__name__)


class TDSReturnsService:

    def __init__(self):
        self.validator  = TDSReturnsValidator()
        self.calculator = TDSReturnsCalculator()
        self.generator  = TDSReturnsExcelGenerator()

    def process(self, file_path: str, company_name: str = '') -> dict:
        return self.process_multiple([file_path], company_name=company_name)

    def process_multiple(self, file_paths, company_name: str = '') -> dict:
        t0 = time.perf_counter()
        logs, errors, warnings = [], [], []

        def log(msg):
            logger.info(msg)
            logs.append({'level': 'INFO', 'message': msg})

        if isinstance(file_paths, str):
            file_paths = [file_paths]

        log(f"Step 1/4: Parsing {len(file_paths)} file(s) — auto-adaptive")
        all_entries = []

        for fp in file_paths:
            fname = fp.rsplit('/', 1)[-1]
            try:
                entries, parse_warnings = parse_for_tds_returns(fp)
                warnings.extend(parse_warnings)
                all_entries.extend(entries)
                log(f"  {fname}: {len(entries)} entries")
            except Exception as exc:
                warnings.append(f"[{fname}] could not be parsed: {exc}")
                logger.warning("Failed: %s: %s", fname, exc)

        log(f"Total: {len(all_entries)} TDS entries from {len(file_paths)} file(s)")

        log("Step 2/4: Validating")
        is_valid, v_errors, v_warnings = self.validator.validate(all_entries)
        errors.extend(v_errors)
        warnings.extend(v_warnings)

        log("Step 3/4: Assigning challans")
        challan_groups = self.calculator.assign_challans(all_entries)
        sections       = self.calculator.compute(all_entries)
        log(f"Generated {len(challan_groups)} challans, {len(sections)} sections")

        if not challan_groups:
            raise EmptyFileError("No TDS data could be extracted from the supplied file(s).")

        log("Step 4/4: Generating Excel")
        excel_bytes = self.generator.generate(
            challan_groups,
            company_name=company_name,
            errors=errors,
            warnings=warnings,
            sections=sections,
        )

        duration_ms = (time.perf_counter() - t0) * 1000
        log(f"Completed in {duration_ms:.0f} ms")

        return {
            'excel_bytes':  excel_bytes,
            'records':      len(all_entries),
            'challans':     len(challan_groups),
            'warnings':     warnings,
            'errors':       errors,
            'logs':         logs,
            'duration_ms':  duration_ms,
        }
