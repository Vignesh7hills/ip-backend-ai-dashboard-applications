"""
POST /api/tds-returns-process

Accepts ONE OR MORE XLSX / XLS / CSV / PDF TDS returns files and returns a
SINGLE merged Excel report (detail + summary + notes sheets).

Multiple uploaded files are parsed independently and merged into one output,
with Challan Serial Numbers running sequentially across all files
(grouped by file / section / month per the Form 26Q requirement).
"""
import os
import tempfile
from typing import List
from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import io

from app.modules.tds_returns.service import TDSReturnsService
from app.utils.file_detector import detect_file_type, validate_file_size
from app.core.exceptions import (
    FileParseError, EmptyFileError, UnsupportedFileTypeError, FinanceBackendError
)
from app.core.config import settings
from app.core.logger import get_logger

logger = get_logger(__name__)
router = APIRouter()
_service = TDSReturnsService()


@router.post(
    "/tds-returns-process",
    summary="Process one or more TDS Returns files -> single merged Excel report",
    tags=["TDS Returns"],
)
async def tds_returns_process(
    files: List[UploadFile] = File(
        ..., description="One or more TDS returns files (XLSX/XLS/CSV/PDF). "
                         "Multiple files are merged into a single output."
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
            fname = up.filename or 'upload.xlsx'
            try:
                validate_file_size(len(content), max_mb=settings.MAX_FILE_SIZE_MB)
                file_type = detect_file_type(fname, content)
            except UnsupportedFileTypeError as exc:
                raise HTTPException(status_code=415, detail=str(exc))

            if file_type not in ('excel', 'csv', 'pdf', 'xml', 'docx'):
                raise HTTPException(
                    status_code=415,
                    detail=f"TDS Returns requires XLSX/XLS/CSV/PDF — '{fname}' is not supported."
                )

            safe = os.path.basename(fname).replace('\\', '_')
            dest = os.path.join(work_dir, safe)
            n = 1
            while os.path.exists(dest):
                root, ext = os.path.splitext(safe)
                dest = os.path.join(work_dir, f"{root}_{n}{ext}")
                n += 1
            with open(dest, 'wb') as fh:
                fh.write(content)
            saved_paths.append(dest)

        logger.info("TDS Returns: %d file(s) received", len(saved_paths))

        result = _service.process_multiple(saved_paths, company_name=company_name)
        excel_bytes: bytes = result['excel_bytes']

        headers = {
            "Content-Disposition": 'attachment; filename="tds_returns_report.xlsx"',
            "X-Records-Processed": str(result['records']),
            "X-Files-Merged": str(len(saved_paths)),
            "X-Challans-Generated": str(result.get('challans', 0)),
            "X-Errors-Count": str(len(result['errors'])),
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
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Unexpected error in TDS returns")
        raise HTTPException(status_code=500, detail={"error": "internal_error", "message": str(exc)})
    finally:
        for p in saved_paths:
            try:
                if os.path.exists(p):
                    os.unlink(p)
            except OSError:
                pass
        try:
            os.rmdir(work_dir)
        except OSError:
            pass
