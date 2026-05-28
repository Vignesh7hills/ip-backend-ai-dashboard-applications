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
        t0 = time.perf_counter()
        logs = []

        def log(msg):
            logger.info(msg)
            logs.append({'level': 'INFO', 'message': msg})

        log("Step 1/4: Parsing TDS file")
        entries = self.parser.parse_file(file_path)
        log(f"Parsed {len(entries)} TDS entries")

        log("Step 2/4: Validating")
        is_valid, errors, warnings = self.validator.validate(entries)

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
