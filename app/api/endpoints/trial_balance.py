"""
POST /api/trial-balance-process

Accepts ONE OR MORE XLSX / XLS / CSV / PDF Balance Sheet, P&L, or Trial Balance files.
Returns a SINGLE merged Excel report (TB Import + Summary + Validation sheets).

Multiple uploaded files (e.g. BS + P&L) are parsed independently and merged into
one balanced Trial Balance output with Closing Stock excluded and Net Profit
transferred to Capital A/c automatically.
"""
import os
import tempfile
from typing import List
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
    summary="Process one or more BS/P&L/TB files → single merged Excel Trial Balance",
    tags=["Trial Balance"],
)
async def trial_balance_process(
    files: List[UploadFile] = File(
        ...,
        description=(
            "One or more Balance Sheet / P&L / Trial Balance files "
            "(XLSX, XLS, CSV, PDF). Multiple files are merged into one output."
        )
    ),
    company_name: str = Form(default=''),
):
    if not files:
        raise HTTPException(status_code=422, detail="No files uploaded.")

    work_dir = tempfile.mkdtemp(dir=str(settings.TEMP_DIR))
    saved_paths: List[str] = []

    try:
        for up in files:
            content = await up.read()
            fname   = up.filename or 'upload.xlsx'
            try:
                validate_file_size(len(content), max_mb=settings.MAX_FILE_SIZE_MB)
                file_type = detect_file_type(fname, content)
            except UnsupportedFileTypeError as exc:
                raise HTTPException(status_code=415, detail=str(exc))

            if file_type not in ('excel', 'csv', 'pdf', 'xml', 'docx'):
                raise HTTPException(
                    status_code=415,
                    detail=f"Trial Balance requires XLSX/XLS/CSV/PDF — '{fname}' is not supported."
                )

            ext  = fname.rsplit('.', 1)[-1].lower()
            safe = os.path.basename(fname).replace('\\\\', '_')
            dest = os.path.join(work_dir, safe)
            n = 1
            while os.path.exists(dest):
                root, ext2 = os.path.splitext(safe)
                dest = os.path.join(work_dir, f"{root}_{n}{ext2}")
                n += 1

            with open(dest, 'wb') as fh:
                fh.write(content)
            saved_paths.append(dest)

        logger.info("Trial Balance: %d file(s) received", len(saved_paths))

        result = _service.process_multiple(saved_paths, company_name=company_name)
        excel_bytes: bytes = result['excel_bytes']

        headers = {
            "Content-Disposition": 'attachment; filename="trial_balance_report.xlsx"',
            "X-Records-Processed": str(result['records']),
            "X-Files-Merged":      str(len(saved_paths)),
            "X-Warnings-Count":    str(len(result['warnings'])),
        }
        return StreamingResponse(
            content=io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )

    except EmptyFileError as exc:
        raise HTTPException(status_code=422, detail={"error": "empty_file",      "message": str(exc)})
    except FileParseError as exc:
        raise HTTPException(status_code=422, detail={"error": "parse_error",     "message": str(exc)})
    except FinanceBackendError as exc:
        raise HTTPException(status_code=422, detail={"error": "processing_error","message": str(exc)})
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Unexpected error in trial balance")
        raise HTTPException(status_code=500, detail={"error": "internal_error",  "message": str(exc)})
    finally:
        for p in saved_paths:
            try:
                if os.path.exists(p): os.unlink(p)
            except OSError:
                pass
        try:
            os.rmdir(work_dir)
        except OSError:
            pass
