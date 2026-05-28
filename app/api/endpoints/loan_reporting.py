"""
POST /api/loan-reporting-process

Accepts a multipart/form-data upload of a PDF ledger (Loans & Advances group).
Returns a downloadable Excel report as a StreamingResponse.
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


@router.post(
    "/loan-reporting-process",
    summary="Process Loans & Advances ledger PDF → Loan Reporting Excel",
    response_description="Excel file (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet)",
    tags=["Loan Reporting"],
)
async def loan_reporting_process(
    file: UploadFile = File(..., description="Loans & Advances PDF ledger"),
    company_name: str = Form(default='', description="Company name for report header"),
):
    """
    Pipeline:
      1. Validate file type (PDF only)
      2. Save to temp file
      3. Parse PDF → extract all account blocks
      4. Normalize transactions
      5. Validate balances
      6. Calculate: taken, repaid, maximum, squared_up
      7. Generate Excel
      8. Return Excel as StreamingResponse (blob)

    Frontend usage:
      ```js
      const formData = new FormData();
      formData.append('file', file);
      axios.post('/api/loan-reporting-process', formData, { responseType: 'blob' })
      ```
    """
    # ── Validation ────────────────────────────────────────────────────────────
    content = await file.read()
    try:
        validate_file_size(len(content), max_mb=settings.MAX_FILE_SIZE_MB)
        file_type = detect_file_type(file.filename or 'upload.pdf')
    except UnsupportedFileTypeError as exc:
        raise HTTPException(status_code=415, detail=str(exc))

    if file_type != 'pdf':
        raise HTTPException(
            status_code=415,
            detail="Loan Reporting requires a PDF file (Tally ledger export)."
        )

    logger.info("Loan Reporting: received file '%s' (%d bytes)", file.filename, len(content))

    # ── Save temp file ────────────────────────────────────────────────────────
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            delete=False, suffix='.pdf', dir=str(settings.TEMP_DIR)
        ) as tmp:
            tmp.write(content)
            tmp_path = tmp.name

        # ── Run pipeline ──────────────────────────────────────────────────────
        result = _service.process(
            pdf_path=tmp_path,
            company_name=company_name or '',
        )

        excel_bytes: bytes = result['excel_bytes']

        # ── Return Excel blob ─────────────────────────────────────────────────
        filename = "loan_reporting_report.xlsx"
        headers = {
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Records-Processed": str(result['records']),
            "X-Warnings-Count": str(len(result['warnings'])),
            "X-Duration-Ms": f"{result['duration_ms']:.0f}",
        }
        return StreamingResponse(
            content=io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )

    except EmptyFileError as exc:
        raise HTTPException(status_code=422, detail={"error": "empty_file", "message": str(exc)})
    except FileParseError as exc:
        raise HTTPException(status_code=422, detail={"error": "parse_error", "message": str(exc)})
    except FinanceBackendError as exc:
        raise HTTPException(status_code=422, detail={"error": "processing_error", "message": str(exc)})
    except Exception as exc:
        logger.exception("Unexpected error in loan reporting")
        raise HTTPException(status_code=500, detail={"error": "internal_error", "message": str(exc)})
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
