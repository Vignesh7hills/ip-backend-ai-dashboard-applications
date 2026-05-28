"""
POST /api/trial-balance-process

Accepts XLSX / XLS / CSV trial balance file.
Returns Excel report with summary + detail sheets.
"""
import os
import tempfile
from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import io

from app.modules.trial_balance.service import TrialBalanceService
from app.utils.file_detector import detect_file_type, validate_file_size
from app.core.exceptions import (
    FileParseError, EmptyFileError, UnsupportedFileTypeError, FinanceBackendError
)
from app.core.config import settings
from app.core.logger import get_logger

logger = get_logger(__name__)
router = APIRouter()
_service = TrialBalanceService()


@router.post(
    "/trial-balance-process",
    summary="Process Trial Balance file → formatted Excel report",
    tags=["Trial Balance"],
)
async def trial_balance_process(
    file: UploadFile = File(..., description="Trial balance XLSX / XLS / CSV"),
    company_name: str = Form(default=''),
):
    content = await file.read()
    try:
        validate_file_size(len(content), max_mb=settings.MAX_FILE_SIZE_MB)
        file_type = detect_file_type(file.filename or 'upload.xlsx')
    except UnsupportedFileTypeError as exc:
        raise HTTPException(status_code=415, detail=str(exc))

    if file_type not in ('excel', 'csv', 'pdf'):
        raise HTTPException(
            status_code=415,
            detail="Trial Balance requires XLSX, XLS, CSV, or PDF file."
        )

    logger.info("Trial Balance: file '%s' (%d bytes)", file.filename, len(content))

    tmp_path = None
    ext = (file.filename or 'upload.xlsx').rsplit('.', 1)[-1].lower()
    try:
        with tempfile.NamedTemporaryFile(
            delete=False, suffix=f'.{ext}', dir=str(settings.TEMP_DIR)
        ) as tmp:
            tmp.write(content)
            tmp_path = tmp.name

        result = _service.process(tmp_path, company_name=company_name)
        excel_bytes: bytes = result['excel_bytes']

        headers = {
            "Content-Disposition": 'attachment; filename="trial_balance_report.xlsx"',
            "X-Records-Processed": str(result['records']),
            "X-Warnings-Count": str(len(result['warnings'])),
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
        logger.exception("Unexpected error in trial balance")
        raise HTTPException(status_code=500, detail={"error": "internal_error", "message": str(exc)})
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
