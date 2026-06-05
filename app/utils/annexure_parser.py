"""
Annexure parser — reads Form 3CD Annexure 31a and 31c XLS files.
Enriches loan report rows with PAN and Address via multi-strategy name matching:
  1. Exact PAN
  2. Exact normalised name
  3. Substring containment
  4. Token overlap (≥2 tokens, ≥60% of shorter name)
  5. Stemmed token overlap (TEXTILE=TEXTILES, etc.)
  6. Character-level similarity via difflib (catches typos like MADANLAL/MANDANLAL)
"""
import re
import difflib
import pandas as pd
from typing import Dict, List
from app.core.logger import get_logger

logger = get_logger(__name__)

_PAN_RE = re.compile(r'^[A-Z]{5}[0-9]{4}[A-Z]$')

# Common word stems for Indian business names
_STEMS = [
    ('TEXTILES', 'TEXTILE'), ('ENTERPRISES', 'ENTERPRISE'),
    ('INDUSTRIES', 'INDUSTRY'), ('TRADERS', 'TRADER'),
    ('EXPORTS', 'EXPORT'), ('IMPORTS', 'IMPORT'),
    ('ASSOCIATES', 'ASSOCIATE'), ('BROTHERS', 'BROTHER'),
    ('SONS', 'SON'), ('AGENCIES', 'AGENCY'),
    ('SYNTHETICS', 'SYNTHETIC'), ('CHEMICALS', 'CHEMICAL'),
    ('MILLS', 'MILL'), ('WORKS', 'WORK'),
    ('SERVICES', 'SERVICE'), ('SOLUTIONS', 'SOLUTION'),
]

def _stem(word: str) -> str:
    """Return the stemmed form of a word for matching."""
    for plural, singular in _STEMS:
        if word == plural:
            return singular
    return word


def _clean(v) -> str:
    s = str(v).strip()
    return '' if s.lower() in ('nan', 'none', '') else s


def _norm_name(name: str) -> str:
    """
    Normalise a name for comparison:
    - Uppercase, remove punctuation
    - Remove type markers: [L], (L), HUF, [Late], (Mohanka), etc.
    - Collapse spaces
    """
    n = name.upper()
    # Remove common suffix markers
    for pat in [
        r'\s*\[\s*L\s*\]', r'\s*\(\s*L\s*\)', r'\s*\[\s*LATE\s*\]',
        r'\s*\(\s*LATE\s*\)', r'\s*HUF\s*', r'\s*\(\s*HUF\s*\)',
        r'\s*\(\s*MOHANKA\s*\)', r'\s*\[L\]', r'\s*\(L\)',
        r'\bL\b',                             # standalone L at end
    ]:
        n = re.sub(pat, '', n, flags=re.I)
    # Remove all punctuation / brackets and collapse spaces
    n = re.sub(r'[^\w\s]', ' ', n)
    n = re.sub(r'\s+', ' ', n).strip()
    return n


def _tokens(norm: str) -> set:
    """Significant tokens (length >= 3)."""
    return set(t for t in norm.split() if len(t) >= 3)


def _stemmed_tokens(norm: str) -> set:
    return set(_stem(t) for t in _tokens(norm))


def _token_overlap_score(a_tokens: set, b_tokens: set) -> float:
    """Return overlap / min(len_a, len_b), or 0 if empty."""
    if not a_tokens or not b_tokens:
        return 0.0
    return len(a_tokens & b_tokens) / min(len(a_tokens), len(b_tokens))


def _char_similarity(a: str, b: str) -> float:
    """Character-level similarity ratio via difflib SequenceMatcher."""
    return difflib.SequenceMatcher(None, a, b).ratio()


def _is_pan(v: str) -> bool:
    return bool(_PAN_RE.match(v.strip().upper())) if v else False


def parse_annexure_files(file_paths: list) -> Dict[str, dict]:
    """
    Parse all annexure XLS files and return a unified reference dict.
    Key = PAN (preferred) else normalised_name.
    Value = {name, address, pan, taken, repaid, maximum, squared_up}
    """
    ref: Dict[str, dict] = {}

    for path in file_paths:
        fname = path.rsplit('/', 1)[-1].lower()
        # Detect annexure type from filename
        is_31a = '31a' in fname
        is_31c = '31c' in fname

        if not (is_31a or is_31c):
            logger.warning("Skipping non-annexure file: %s", fname)
            continue

        try:
            df = pd.read_excel(path, header=None, engine='openpyxl')
        except Exception as e:
            logger.warning("Cannot read annexure %s: %s", path, e)
            continue

        # ── Find header row ───────────────────────────────────────────────────
        hdr_row = None
        for ri in range(min(5, len(df))):
            row_text = ' '.join(
                str(v).lower() for v in df.iloc[ri]
                if str(v).strip() not in ('nan', '')
            )
            if 'name' in row_text and ('lender' in row_text or 'payee' in row_text):
                hdr_row = ri
                break

        if hdr_row is None:
            logger.warning("No header row found in %s", fname)
            continue

        hdr = [str(df.iloc[hdr_row, ci]).lower() for ci in range(df.shape[1])]

        def find_col(keywords):
            for ci, h in enumerate(hdr):
                if any(kw in h for kw in keywords):
                    return ci
            return -1

        col_sn   = find_col(['sn', 'sr'])
        col_name = find_col(['name of the lender', 'name of payee', 'name'])
        col_addr = find_col(['address'])
        col_pan  = find_col(['pan'])
        col_amt  = find_col(['amount of loan', 'amount of the repayment', 'amount of repayment'])
        col_max  = find_col(['maximum amount'])
        col_sq   = find_col(['squared up'])

        if col_name < 0:
            logger.warning("Cannot find name column in %s", fname)
            continue

        logger.info("Parsing %s (31a=%s 31c=%s): name=%d addr=%d pan=%d amt=%d max=%d",
                    fname, is_31a, is_31c, col_name, col_addr, col_pan, col_amt, col_max)

        for ri in range(hdr_row + 1, len(df)):
            row = df.iloc[ri]

            def sv(ci):
                return _clean(row[ci]) if 0 <= ci < len(row) else ''

            sn_val = sv(col_sn).replace('.0', '').strip()
            if not sn_val or not re.match(r'^\d+$', sn_val):
                continue

            name = sv(col_name)
            if not name or len(name) < 2:
                continue

            addr = sv(col_addr)
            pan  = sv(col_pan).upper().strip() if col_pan >= 0 else ''

            def parse_amt(v):
                try:
                    return float(str(v).replace(',', '').strip() or '0')
                except ValueError:
                    return 0.0

            amount  = parse_amt(sv(col_amt))
            maximum = parse_amt(sv(col_max))
            sq_up   = 'YES' if sv(col_sq) and 'yes' in sv(col_sq).lower() else 'NO'

            # Choose lookup key: PAN > normalised name
            if _is_pan(pan):
                key = pan.upper()
            else:
                pan = ''
                key = _norm_name(name)

            if key not in ref:
                ref[key] = {
                    'name': name, 'address': addr, 'pan': pan,
                    'taken': 0.0, 'repaid': 0.0,
                    'maximum': maximum, 'squared_up': sq_up,
                }

            if is_31a:
                ref[key]['taken'] += amount
            elif is_31c:
                ref[key]['repaid'] += amount

            ref[key]['maximum'] = max(ref[key]['maximum'], maximum)
            if addr and len(addr) > len(ref[key]['address']):
                ref[key]['address'] = addr
            if pan and not ref[key]['pan']:
                ref[key]['pan'] = pan
                # Re-key by PAN
                if key != pan.upper():
                    ref[pan.upper()] = ref.pop(key)

    logger.info("Annexure parsed: %d reference records", len(ref))
    return ref


def enrich_with_annexure(report_rows: list, annexure_ref: Dict[str, dict]) -> list:
    """
    Enrich loan report rows with PAN and Address from annexure.

    Matching order (stops at first hit):
      1. Exact PAN
      2. Exact normalised name
      3. Normalised substring containment
      4. Token overlap ≥ 60% (regular tokens)
      5. Stemmed token overlap ≥ 60% (handles TEXTILE/TEXTILES etc.)
      6. Character similarity ≥ 0.82 on the full normalised name (typo tolerance)
    """
    if not annexure_ref:
        return report_rows

    # Pre-build indices for speed
    name_index:    Dict[str, str] = {}   # norm_name → key
    stem_index:    Dict[str, str] = {}   # frozenset(stemmed_tokens) → key (first match)

    for key, rec in annexure_ref.items():
        nn = _norm_name(rec['name'])
        name_index[nn] = key
        sk = frozenset(_stemmed_tokens(nn))
        if sk not in stem_index:
            stem_index[sk] = key

    for row in report_rows:
        row_name = getattr(row, 'name', '') or ''
        row_pan  = (getattr(row, 'pan', '') or '').strip().upper()
        match_key = None

        nn = _norm_name(row_name)
        nn_tok  = _tokens(nn)
        nn_stem = _stemmed_tokens(nn)

        # ── 1. Exact PAN ──────────────────────────────────────────────────────
        if not match_key and row_pan and _is_pan(row_pan) and row_pan in annexure_ref:
            match_key = row_pan

        # ── 2. Exact normalised name ──────────────────────────────────────────
        if not match_key and nn in name_index:
            match_key = name_index[nn]

        # ── 3. Substring containment ──────────────────────────────────────────
        if not match_key:
            for idx_name, key in name_index.items():
                if len(nn) >= 4 and len(idx_name) >= 4:
                    if nn in idx_name or idx_name in nn:
                        match_key = key
                        break

        # ── 4. Regular token overlap ≥ 60% ───────────────────────────────────
        if not match_key and len(nn_tok) >= 2:
            best_score, best_key = 0.0, None
            for idx_name, key in name_index.items():
                idx_tok = _tokens(idx_name)
                score = _token_overlap_score(nn_tok, idx_tok)
                if score >= 0.60 and score > best_score:
                    best_score, best_key = score, key
            if best_key:
                match_key = best_key

        # ── 5. Stemmed token overlap ≥ 60% ───────────────────────────────────
        if not match_key and len(nn_stem) >= 2:
            best_score, best_key = 0.0, None
            for idx_name, key in name_index.items():
                idx_stem = _stemmed_tokens(idx_name)
                score = _token_overlap_score(nn_stem, idx_stem)
                if score >= 0.60 and score > best_score:
                    best_score, best_key = score, key
            if best_key:
                match_key = best_key

        # ── 6. Character-level similarity ≥ 0.82 on full name ────────────────
        if not match_key:
            best_score, best_key = 0.0, None
            for idx_name, key in name_index.items():
                sim = _char_similarity(nn, idx_name)
                if sim >= 0.82 and sim > best_score:
                    best_score, best_key = sim, key
            if best_key:
                match_key = best_key

        # ── 7. First+Last word char-sim ≥ 0.88 (catches middle-name typos) ──
        if not match_key:
            nn_fl = _first_last(nn)
            best_score, best_key = 0.0, None
            for idx_name, key in name_index.items():
                idx_fl = _first_last(idx_name)
                sim = _char_similarity(nn_fl, idx_fl)
                if sim >= 0.88 and sim > best_score:
                    best_score, best_key = sim, key
            if best_key:
                match_key = best_key

        if match_key and match_key in annexure_ref:
            rec = annexure_ref[match_key]
            if not row_pan and rec['pan']:
                row.pan = rec['pan']
            cur_addr = getattr(row, 'address', '') or ''
            ann_addr = rec['address'] or ''
            if ann_addr and len(ann_addr) > len(cur_addr):
                row.address = ann_addr
            logger.debug("Enriched '%s' → key=%s pan=%s", row_name, match_key, row.pan)

    return report_rows


def _first_last(norm: str) -> str:
    """Return first + last word of a normalised name."""
    parts = norm.split()
    return f"{parts[0]} {parts[-1]}" if len(parts) >= 2 else norm
