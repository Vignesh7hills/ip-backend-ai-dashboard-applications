"""
POST /api/loan-reporting-process

Accepts ANY file format: PDF, Excel (xlsx/xls/xlsm/ods), CSV, TSV, TXT.
Returns a downloadable Excel report as a StreamingResponse.

Auto-detects file format using magic bytes + extension.
Falls back between parsers automatically.

Frontend usage:
  ```js
  const formData = new FormData();
  formData.append('file', file);
  axios.post('/api/loan-reporting-process', formData, { responseType: 'blob' })
  ```
"""
import os
import tempfile
from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import io

from app.modules.loan_reporting.service import LoanReportingService
from app.utils.file_detector import detect_file_type, validate_file_size
from app.core.exceptions import (
    FileParseError, EmptyFileError, UnsupportedFileTypeError, FinanceBackendError
)
from app.core.config import settings
from app.core.logger import get_logger

logger = get_logger(__name__)
router = APIRouter()
_service = LoanReportingService()

# All accepted MIME types / extensions
_ACCEPTED_TYPES = {'pdf', 'excel', 'csv', 'xml', 'docx'}


@router.post(
    "/loan-reporting-process",
    summary="Process Loans & Advances ledger (any format) → Loan Reporting Excel",
    response_description="Excel file (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet)",
    tags=["Loan Reporting"],
)
async def loan_reporting_process(
    file: UploadFile = File(..., description=(
        "Loans & Advances ledger in ANY format: "
        "PDF (Tally export), Excel (.xlsx/.xls/.xlsm/.ods), CSV, TSV, or TXT"
    )),
    company_name: str = Form(default='', description="Company name for report header"),
):
    """
    Pipeline:
      1. Accept any file format — magic-byte detection, not just extension
      2. Auto-select parser (PDF or Excel/CSV)
      3. Parse → extract all account blocks
      4. Normalize transactions
      5. Validate balances
      6. Calculate: taken, repaid, maximum, squared_up
      7. Generate Excel
      8. Return Excel as StreamingResponse (blob)
    """
    content = await file.read()
    fname = file.filename or 'upload'

    # ── Size validation ───────────────────────────────────────────────────────
    try:
        validate_file_size(len(content), max_mb=settings.MAX_FILE_SIZE_MB)
    except UnsupportedFileTypeError as exc:
        raise HTTPException(status_code=413, detail=str(exc))

    # ── File-type detection (magic bytes + extension) ─────────────────────────
    try:
        file_type = detect_file_type(fname, content)
    except UnsupportedFileTypeError as exc:
        raise HTTPException(status_code=415, detail=str(exc))

    logger.info(
        "Loan Reporting: file='%s' type=%s size=%d bytes",
        fname, file_type, len(content)
    )

    # ── Determine file extension for saving temp file ────────────────────────
    ext_map = {
        'pdf':   '.pdf',
        'excel': _get_excel_ext(fname),
        'csv':   _get_csv_ext(fname),
        'xml':   '.xml',
        'docx':  '.docx',
    }
    suffix = ext_map.get(file_type, os.path.splitext(fname)[1] or '.tmp')

    # ── Save temp file ────────────────────────────────────────────────────────
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            delete=False, suffix=suffix, dir=str(settings.TEMP_DIR)
        ) as tmp:
            tmp.write(content)
            tmp_path = tmp.name

        # ── Run pipeline ──────────────────────────────────────────────────────
        result = _service.process(
            pdf_path=tmp_path,
            company_name=company_name or '',
            file_type=file_type,
        )

        excel_bytes: bytes = result['excel_bytes']

        # ── Return Excel blob ─────────────────────────────────────────────────
        out_filename = "loan_reporting_report.xlsx"
        headers = {
            "Content-Disposition": f'attachment; filename="{out_filename}"',
            "X-Records-Processed": str(result['records']),
            "X-Warnings-Count":    str(len(result['warnings'])),
            "X-Duration-Ms":       f"{result['duration_ms']:.0f}",
            "X-Source-Format":     file_type,
        }
        return StreamingResponse(
            content=io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )

    except EmptyFileError as exc:
        raise HTTPException(status_code=422, detail={"error": "empty_file",       "message": str(exc)})
    except FileParseError as exc:
        raise HTTPException(status_code=422, detail={"error": "parse_error",      "message": str(exc)})
    except FinanceBackendError as exc:
        raise HTTPException(status_code=422, detail={"error": "processing_error", "message": str(exc)})
    except Exception as exc:
        logger.exception("Unexpected error in loan reporting")
        raise HTTPException(status_code=500, detail={"error": "internal_error",   "message": str(exc)})
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


def _get_excel_ext(fname: str) -> str:
    ext = os.path.splitext(fname)[1].lower()
    return ext if ext in ('.xlsx', '.xls', '.xlsm', '.ods') else '.xlsx'


def _get_csv_ext(fname: str) -> str:
    ext = os.path.splitext(fname)[1].lower()
    return ext if ext in ('.csv', '.tsv', '.txt') else '.csv'
