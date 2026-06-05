"""
Universal Content Detector — AUTO-ADAPTIVE for all 3 modules.

Reads a file (any format) and detects WHAT KIND OF DATA it contains,
independent of which module it was uploaded to.

Content types:
  'ledger'   — Date + Dr/Cr/Balance transactions (Loan Reporting)
  'tds'      — Party/PAN + Amount + TDS/Rate (TDS Returns)
  'tb'       — Account + Debit + Credit columns (Trial Balance)
  'bs'       — Liabilities/Assets two-sided (Trial Balance BS format)
  'unknown'  — fallback, try all parsers

Used by every module's service to route to the correct parser
even when the wrong module is selected by the user.
"""

import re
import pandas as pd
from typing import Dict, Tuple
from app.core.logger import get_logger

logger = get_logger(__name__)

# Signature keywords per content type
_LEDGER_DATE   = re.compile(r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}')
_LEDGER_BAL    = re.compile(r'balance|bal\b', re.I)
_LEDGER_DRCRBAL = re.compile(r'\b(debit|credit|dr|cr|dr\.|cr\.)\b', re.I)

_TDS_KEYS = [
    'pan', 'pan no', 'pan no.', 'permanent account',
    'tds', 'tds %', 'tds amt', 'tax deducted',
    'challan', 'section', 'deductee',
    'cr.  amount', 'cr. amount', 'applicable amt',
]
_TDS_REQUIRED = ['pan', 'tds', 'challan', 'deductee', 'section', 'rate']

_TB_KEYS = [
    'amount (dr)', 'amount (cr)', 'amount(dr)', 'amount(cr)',
    'debit', 'credit', 'ledger name', 'account name',
    'opening dr', 'closing cr', 'trial balance',
]
_BS_KEYS = ['liabilit', 'assets', 'capital', 'p&l', 'profit', 'loss', 'trading']

_LOAN_KEYS = [
    'voucher no', 'voucher', 'narration', 'loans', 'advances',
    'loan & advance', 'ledger a/c',
]


def detect_content_type(file_path: str) -> str:
    """
    Read the first sheet/page of a file and detect its content type.
    Returns: 'ledger' | 'tds' | 'tb' | 'bs' | 'unknown'
    """
    try:
        df = _read_sample(file_path)
        if df is None or df.empty:
            return 'unknown'
        return _classify(df, file_path)
    except Exception as e:
        logger.warning("Content detection failed for %s: %s", file_path, e)
        return 'unknown'


def _read_sample(file_path: str) -> pd.DataFrame:
    """Read first 50 rows of the file into a DataFrame."""
    ext = file_path.rsplit('.', 1)[-1].lower()
    try:
        if ext in ('xlsx', 'xlsm'):
            return pd.read_excel(file_path, header=None, dtype=str,
                                 engine='openpyxl', nrows=50)
        elif ext == 'xls':
            return pd.read_excel(file_path, header=None, dtype=str,
                                 engine='xlrd', nrows=50)
        elif ext == 'ods':
            return pd.read_excel(file_path, header=None, dtype=str,
                                 engine='odf', nrows=50)
        elif ext in ('csv', 'tsv', 'txt'):
            sep = '\t' if ext == 'tsv' else None
            return pd.read_csv(file_path, header=None, dtype=str,
                               sep=sep, engine='python', nrows=50,
                               on_bad_lines='skip', encoding_errors='replace')
        elif ext == 'pdf':
            return _pdf_sample(file_path)
    except Exception:
        pass
    return None


def _pdf_sample(file_path: str) -> pd.DataFrame:
    """Extract first ~50 lines from PDF as single-column DataFrame."""
    lines = []
    try:
        from pypdf import PdfReader
        reader = PdfReader(file_path)
        for page in reader.pages[:3]:
            for line in (page.extract_text() or '').split('\n'):
                line = line.strip()
                if line:
                    lines.append([line])
                if len(lines) >= 50:
                    break
    except Exception:
        pass
    if not lines:
        return None
    return pd.DataFrame(lines, dtype=str)


def _classify(df: pd.DataFrame, file_path: str) -> str:
    """Score the DataFrame against each content type and return best match."""
    # Flatten all cell text to lowercase for scanning
    all_text = ' '.join(
        str(c).lower().strip()
        for row in df.values
        for c in row
        if str(c).strip() not in ('', 'nan', 'none')
    )

    scores: Dict[str, int] = {'ledger': 0, 'tds': 0, 'tb': 0, 'bs': 0}

    # ── TDS signals ───────────────────────────────────────────────────────────
    for kw in _TDS_KEYS:
        if kw in all_text:
            scores['tds'] += 2
    # Strong TDS signal: PAN + (TDS or challan)
    if 'pan' in all_text and ('tds' in all_text or 'challan' in all_text):
        scores['tds'] += 5

    # ── Trial Balance signals ─────────────────────────────────────────────────
    for kw in _TB_KEYS:
        if kw in all_text:
            scores['tb'] += 2
    if 'amount (dr)' in all_text and 'amount (cr)' in all_text:
        scores['tb'] += 8   # very strong signal
    if 'debit' in all_text and 'credit' in all_text and 'ledger' in all_text:
        scores['tb'] += 5

    # ── Balance Sheet signals ─────────────────────────────────────────────────
    for kw in _BS_KEYS:
        if kw in all_text:
            scores['bs'] += 1
    if 'liabilit' in all_text and 'asset' in all_text:
        scores['bs'] += 8
    if scores['bs'] >= 4:
        scores['tb'] += scores['bs']   # BS is a sub-type of TB

    # ── Ledger/Loan signals ───────────────────────────────────────────────────
    # Count actual date values in the DataFrame
    date_count = sum(
        1 for row in df.values for c in row
        if _LEDGER_DATE.search(str(c))
    )
    if date_count >= 3:
        scores['ledger'] += date_count  # many dates → transaction ledger
    for kw in _LOAN_KEYS:
        if kw in all_text:
            scores['ledger'] += 2
    if 'balance' in all_text and date_count >= 2:
        scores['ledger'] += 3
    if ('debit' in all_text or ' dr ' in all_text) and date_count >= 2:
        scores['ledger'] += 3

    # ── Filename hints ────────────────────────────────────────────────────────
    fname = file_path.rsplit('/', 1)[-1].lower()
    if any(k in fname for k in ('tds', '94', '26q', 'challan', 'deductee')):
        scores['tds'] += 4
    if any(k in fname for k in ('trial', 'balance', 'tb ', 'bs ', 'p&l', 'pl ')):
        scores['tb'] += 4
    if any(k in fname for k in ('ledger', 'loan', 'advance', 'borrowing')):
        scores['ledger'] += 4

    best = max(scores, key=lambda k: scores[k])
    best_score = scores[best]

    logger.info("Content detection scores for %s: %s → %s",
                file_path.rsplit('/', 1)[-1], scores, best)

    if best_score < 2:
        return 'unknown'
    return best
