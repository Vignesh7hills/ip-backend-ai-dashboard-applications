"""Trial Balance Service — parse → validate → calculate → generate Excel."""
import time
from app.modules.trial_balance.parser import TrialBalanceParser
from app.modules.trial_balance.validator import TrialBalanceValidator
from app.modules.trial_balance.calculator import TrialBalanceCalculator
from app.modules.trial_balance.generator import TrialBalanceExcelGenerator
from app.core.logger import get_logger
from app.core.exceptions import EmptyFileError

logger = get_logger(__name__)


class TrialBalanceService:

    def __init__(self):
        self.parser    = TrialBalanceParser()
        self.validator = TrialBalanceValidator()
        self.calculator = TrialBalanceCalculator()
        self.generator = TrialBalanceExcelGenerator()

    def process(self, file_path: str, company_name: str = '') -> dict:
        t0 = time.perf_counter()
        logs = []

        def log(msg):
            logger.info(msg)
            logs.append({'level': 'INFO', 'message': msg})

        log("Step 1/4: Parsing trial balance file")
        entries = self.parser.parse_file(file_path)
        log(f"Parsed {len(entries)} entries")

        log("Step 2/4: Validating")
        is_valid, errors, warnings = self.validator.validate(entries)
        if not is_valid:
            for e in errors:
                logger.warning("Validation error: %s", e)

        log("Step 3/4: Calculating group totals")
        groups = self.calculator.compute(entries)
        log(f"Computed {len(groups)} groups")

        if not groups:
            raise EmptyFileError("No data groups found in trial balance.")

        log("Step 4/4: Generating Excel")
        excel_bytes = self.generator.generate(
            groups,
            entries_flat=entries,
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
