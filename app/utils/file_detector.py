"""
File-type detection utilities.
"""
import os
from pathlib import Path
from app.core.exceptions import UnsupportedFileTypeError

SUPPORTED = {
    ".pdf": "pdf",
    ".xlsx": "excel",
    ".xls": "excel",
    ".csv": "csv",
    ".xml": "xml",
}


def detect_file_type(filename: str) -> str:
    """Return a normalised file-type string, or raise UnsupportedFileTypeError."""
    ext = Path(filename).suffix.lower()
    if ext not in SUPPORTED:
        raise UnsupportedFileTypeError(
            f"File type '{ext}' is not supported.",
            {"filename": filename, "allowed": list(SUPPORTED.keys())},
        )
    return SUPPORTED[ext]


def validate_file_size(size_bytes: int, max_mb: int = 50) -> None:
    max_bytes = max_mb * 1024 * 1024
    if size_bytes > max_bytes:
        raise UnsupportedFileTypeError(
            f"File size {size_bytes / 1024 / 1024:.1f} MB exceeds limit of {max_mb} MB."
        )
