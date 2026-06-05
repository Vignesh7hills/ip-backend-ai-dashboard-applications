"""
POST /api/loan-reporting-process

AUTO-ADAPTIVE — accepts any mix of files in a single upload field:
  • Ledger files (PDF, Excel, CSV) — any filename → loan report data
  • Annexure files — auto-detected by filename containing 31a/31c/annexure

All files go into ONE form field called 'files' (or legacy 'file').
No separate 'annexures' field needed — backend auto-separates them.

Frontend usage (unchanged from old single-file upload):
  const formData = new FormData();
  formData.append('files', ledgerPDF);
  formData.append('files', annexure31a);   ← just add them all together
  formData.append('files', annexure31c);
  axios.post('/api/loan-reporting-process', formData, { responseType: 'blob' })
"""
import os
import re
import tempfile
from typing import List, Optional
from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import io

from app.modules.loan_reporting.service import LoanReportingService
from app.utils.file_detector import detect_file_type, validate_file_size
from app.utils.annexure_parser import parse_annexure_files
from app.core.exceptions import (
    FileParseError, EmptyFileError, UnsupportedFileTypeError, FinanceBackendError
)
from app.core.config import settings
from app.core.logger import get_logger

logger = get_logger(__name__)
router = APIRouter()
_service = LoanReportingService()

# Keywords that flag a file as an annexure reference document
_ANNEXURE_KW = re.compile(
    r'(31\s*[ac]|annexure|annex|annx|ann_)',
    re.IGNORECASE
)


def _is_annexure_file(filename: str) -> bool:
    """Return True if filename suggests this is an Annexure 31a/31c file."""
    return bool(_ANNEXURE_KW.search(filename))


@router.post(
    "/loan-reporting-process",
    summary="Process Loans & Advances (multiple files + auto-detected annexures) → Excel",
    tags=["Loan Reporting"],
)
async def loan_reporting_process(
    files: List[UploadFile] = File(..., description=(
        "All files together: ledger PDFs/Excel/CSV AND optional Annexure 31a/31c XLS. "
        "Annexures are auto-detected from filename (must contain 31a, 31c, or 'annexure')."
    )),
    # Legacy single-file field — kept for backward compatibility
    file: Optional[UploadFile] = File(default=None),
    company_name: str = Form(default=''),
):
    """
    Auto-separates uploaded files into ledger files and annexure files by filename.
    Then: parse ledgers → enrich PAN/Address from annexures → generate Excel.
    """
    # Merge legacy 'file' field with 'files'
    all_uploads: List[UploadFile] = list(files or [])
    if file is not None:
        all_uploads.insert(0, file)

    if not all_uploads:
        raise HTTPException(status_code=422, detail="No files provided.")

    ledger_paths: List[str]  = []
    annex_paths:  List[str]  = []

    try:
        for upload in all_uploads:
            content = await upload.read()
            fname   = upload.filename or 'upload'

            try:
                validate_file_size(len(content), max_mb=settings.MAX_FILE_SIZE_MB)
            except UnsupportedFileTypeError as exc:
                raise HTTPException(status_code=413, detail=str(exc))

            # ── Auto-detect: annexure or ledger? ─────────────────────────────
            if _is_annexure_file(fname):
                ext = os.path.splitext(fname)[1] or '.xls'
                with tempfile.NamedTemporaryFile(
                    delete=False, suffix=ext, dir=str(settings.TEMP_DIR)
                ) as tmp:
                    tmp.write(content)
                    tmp_path = tmp.name

                # Rename so 31a/31c is preserved in path for parser detection
                final_path = tmp_path + '_' + re.sub(r'[/\\]', '_', fname)
                os.rename(tmp_path, final_path)
                annex_paths.append(final_path)
                logger.info("Annexure file detected: %s (%d bytes)", fname, len(content))

            else:
                # Ledger file
                try:
                    file_type = detect_file_type(fname, content)
                except UnsupportedFileTypeError as exc:
                    raise HTTPException(status_code=415, detail=str(exc))

                ext = _ext_for(fname, file_type)
                with tempfile.NamedTemporaryFile(
                    delete=False, suffix=ext, dir=str(settings.TEMP_DIR)
                ) as tmp:
                    tmp.write(content)
                    ledger_paths.append(tmp.name)
                logger.info("Ledger file: %s → %s (%d bytes)", fname, file_type, len(content))

        if not ledger_paths:
            raise HTTPException(status_code=422, detail={
                "error": "no_ledger",
                "message": (
                    "No ledger files found. All files were detected as annexures. "
                    "Please include at least one loan ledger PDF or Excel file."
                )
            })

        # ── Parse annexure reference ──────────────────────────────────────────
        annexure_ref = {}
        if annex_paths:
            try:
                annexure_ref = parse_annexure_files(annex_paths)
                logger.info("Annexure ref built: %d records from %d file(s)",
                            len(annexure_ref), len(annex_paths))
            except Exception as e:
                logger.warning("Annexure parse warning (non-fatal): %s", e)
        else:
            logger.info("No annexure files uploaded — PAN/Address will be blank")

        # ── Run pipeline ──────────────────────────────────────────────────────
        result = _service.process_multiple(
            file_paths=ledger_paths,
            company_name=company_name or '',
            annexure_ref=annexure_ref,
        )

        excel_bytes: bytes = result['excel_bytes']

        ann_note = f"{len(annex_paths)} annexure(s) used" if annex_paths else "no annexures"
        headers = {
            "Content-Disposition": 'attachment; filename="loan_reporting_report.xlsx"',
            "X-Records-Processed": str(result['records']),
            "X-Ledger-Files":      str(len(ledger_paths)),
            "X-Annexure-Files":    str(len(annex_paths)),
            "X-Annexure-Records":  str(len(annexure_ref)),
            "X-Warnings-Count":    str(len(result.get('warnings', []))),
            "X-Duration-Ms":       f"{result.get('duration_ms', 0):.0f}",
        }
        logger.info("Loan Reporting done: %d records, %s",
                    result['records'], ann_note)

        return StreamingResponse(
            content=io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )

    except EmptyFileError as exc:
        raise HTTPException(status_code=422, detail={"error": "empty_file",    "message": str(exc)})
    except FileParseError as exc:
        raise HTTPException(status_code=422, detail={"error": "parse_error",   "message": str(exc)})
    except FinanceBackendError as exc:
        raise HTTPException(status_code=422, detail={"error": "processing",    "message": str(exc)})
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Unexpected error in loan reporting")
        raise HTTPException(status_code=500, detail={"error": "internal_error","message": str(exc)})
    finally:
        for p in ledger_paths + annex_paths:
            try:
                if os.path.exists(p): os.unlink(p)
            except Exception:
                pass


def _ext_for(fname: str, file_type: str) -> str:
    ext = os.path.splitext(fname)[1].lower()
    if ext: return ext
    return {
        'pdf': '.pdf', 'excel': '.xlsx', 'csv': '.csv',
        'xml': '.xml', 'docx': '.docx',
    }.get(file_type, '.tmp')
