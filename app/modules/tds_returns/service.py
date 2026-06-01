"""TDS Returns Service — full pipeline."""
import time
from app.modules.tds_returns.parser import TDSReturnsParser
from app.modules.tds_returns.validator import TDSReturnsValidator
from app.modules.tds_returns.calculator import TDSReturnsCalculator
from app.modules.tds_returns.generator import TDSReturnsExcelGenerator
from app.core.logger import get_logger
from app.core.exceptions import EmptyFileError

logger = get_logger(__name__)


class TDSReturnsService:

    def __init__(self):
        self.parser    = TDSReturnsParser()
        self.validator = TDSReturnsValidator()
        self.calculator = TDSReturnsCalculator()
        self.generator = TDSReturnsExcelGenerator()

    def process(self, file_path: str, company_name: str = '') -> dict:
        """Single-file convenience wrapper around process_multiple."""
        return self.process_multiple([file_path], company_name=company_name)

    def process_multiple(self, file_paths, company_name: str = '') -> dict:
        """
        Parse one OR many TDS input files and produce a SINGLE merged report.

        Each file is parsed independently; entries from every file are tagged
        with their own source_file / source_group (done by the parser). The
        combined entry list is then validated, challan-numbered, and rendered
        once — so multiple uploads collapse into one output workbook with
        challan serials running sequentially across all files.
        """
        t0 = time.perf_counter()
        logs = []
        errors = []
        warnings = []

        def log(msg):
            logger.info(msg)
            logs.append({'level': 'INFO', 'message': msg})

        if isinstance(file_paths, str):
            file_paths = [file_paths]

        log(f"Step 1/4: Parsing {len(file_paths)} TDS file(s)")
        entries = []
        per_file = []
        for fp in file_paths:
            fname = fp.rsplit('/', 1)[-1]
            try:
                e = self.parser.parse_file(fp)
            except Exception as exc:
                warnings.append(f"[{fname}] could not be parsed: {exc}")
                logger.warning("Failed to parse %s: %s", fname, exc)
                continue
            if not e:
                warnings.append(f"[{fname}] no TDS rows detected")
            entries.extend(e)
            per_file.append((fname, len(e)))

        for fname, n in per_file:
            log(f"  {fname}: {n} entries")
        log(f"Parsed {len(entries)} TDS entries from {len(per_file)} file(s)")

        log("Step 2/4: Validating")
        is_valid, v_errors, v_warnings = self.validator.validate(entries)
        errors.extend(v_errors)
        warnings.extend(v_warnings)

        log("Step 3/4: Assigning challans & computing summaries")
        challan_groups = self.calculator.assign_challans(entries)
        sections       = self.calculator.compute(entries)
        log(f"Generated {len(challan_groups)} challans, {len(sections)} sections")

        if not challan_groups:
            raise EmptyFileError("No TDS data found.")

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
            'excel_bytes':   excel_bytes,
            'records':       len(entries),
            'challans':      len(challan_groups),
            'warnings':      warnings,
            'errors':        errors,
            'logs':          logs,
            'duration_ms':   duration_ms,
        }
