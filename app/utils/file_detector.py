"""
File-type detection utilities — AUTO-ADAPTIVE.

Detects file type using BOTH:
  1. Magic bytes (first 8 bytes of file content) — reliable regardless of extension
  2. File extension — fallback / confirmation

Supports: PDF, Excel (xlsx/xls/xlsm/ods), CSV, XML, Word DOCX, plain text.
Unknown but safely parseable files are attempted as CSV.
"""
import os
from pathlib import Path
from app.core.exceptions import UnsupportedFileTypeError

# Known magic byte signatures
_MAGIC = {
    b'\x50\x4b\x03\x04': 'zip_based',   # ZIP (xlsx, xlsm, docx, odt, ods)
    b'\xd0\xcf\x11\xe0': 'ole',          # OLE (xls, doc, ppt — legacy Office)
    b'\x25\x50\x44\x46': 'pdf',          # PDF: %PDF
    b'\xef\xbb\xbf':     'utf8bom',      # UTF-8 BOM (text/csv)
    b'\xff\xfe':         'utf16le',      # UTF-16 LE BOM (text)
    b'\xfe\xff':         'utf16be',      # UTF-16 BE BOM (text)
    b'<?xm':             'xml',          # XML
    b'<htm':             'html',
}

_EXT_MAP = {
    '.pdf':  'pdf',
    '.xlsx': 'excel',
    '.xlsm': 'excel',
    '.xls':  'excel',
    '.ods':  'excel',
    '.csv':  'csv',
    '.tsv':  'csv',
    '.txt':  'csv',      # treat as CSV (tab/comma separated text)
    '.xml':  'xml',
    '.json': 'json',
    '.docx': 'docx',
    '.doc':  'docx',
}

ACCEPTED_TYPES = {'pdf', 'excel', 'csv', 'xml', 'docx'}


def detect_file_type(filename: str, content: bytes = b'') -> str:
    """
    Return a normalised file-type string.
    Uses magic bytes when content is provided, falls back to extension.
    Raises UnsupportedFileTypeError only for clearly binary/unknown content
    that cannot be treated as text.
    """
    ext = Path(filename).suffix.lower()

    # ── Magic-byte detection (most reliable) ─────────────────────────────────
    if content:
        magic = _detect_magic(content, ext)
        if magic:
            return magic

    # ── Extension fallback ────────────────────────────────────────────────────
    if ext in _EXT_MAP:
        return _EXT_MAP[ext]

    # ── Sniff as text/csv (last resort for .txt, no-ext, unknown) ────────────
    if content:
        try:
            sample = content[:4096].decode('utf-8', errors='strict')
            # If decodable as UTF-8 and contains comma/tab/newline → treat as CSV
            if any(c in sample for c in (',', '\t', '\n')):
                return 'csv'
        except UnicodeDecodeError:
            pass

    # ── Never hard-fail on format: default by extension family, otherwise
    #    assume Excel so the parser (which has multi-engine fallbacks) can try. ─
    if ext in _EXT_MAP:
        return _EXT_MAP[ext]
    return 'excel'


def _detect_magic(content: bytes, ext: str) -> str:
    """Detect file type from magic bytes. Returns type string or ''."""
    head = content[:8]

    # PDF
    if head[:4] == b'%PDF':
        return 'pdf'

    # OLE2 compound document → legacy xls / doc
    if head[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        if ext == '.doc':
            return 'docx'
        return 'excel'   # xls default

    # ZIP-based (xlsx, xlsm, docx, ods, odt)
    if head[:4] == b'PK\x03\x04':
        if ext in ('.docx', '.doc', '.odt'):
            return 'docx'
        if ext in ('.ods',):
            return 'excel'
        # default: excel (xlsx/xlsm)
        return 'excel'

    # XML / HTML
    if head[:5] in (b'<?xml', b'<html', b'<!DOC'):
        return 'xml'
    if head[:4] == b'<?xm':
        return 'xml'

    # UTF-8 BOM → likely CSV/text
    if head[:3] == b'\xef\xbb\xbf':
        return 'csv'

    # UTF-16 BOM → likely CSV/text
    if head[:2] in (b'\xff\xfe', b'\xfe\xff'):
        return 'csv'

    return ''


def validate_file_size(size_bytes: int, max_mb: int = 50) -> None:
    max_bytes = max_mb * 1024 * 1024
    if size_bytes > max_bytes:
        raise UnsupportedFileTypeError(
            f"File size {size_bytes / 1024 / 1024:.1f} MB exceeds limit of {max_mb} MB."
        )
