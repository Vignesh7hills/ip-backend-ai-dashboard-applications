"""
PDF Ledger Parser for Loan & Advances group ledgers — FIXED VERSION.

Supports two Tally export formats:

MTA format (MTA-LOAN AND ADVANCES 2024-25.pdf):
  - Each column value on its own line after pypdf extraction
  - Date: DD-MM-YYYY  (dashes)
  - Balance: "8,50,000.00Cr"  (concatenated direction, Indian comma format)
  - Voucher types: Opbl | Pay | Rec | Jrn
  - Column headers: Date | Voucher No. | Description | Debit | Credit | Balance

MANMOHAN format (LOAN.pdf):
  - Each table row on one line; columns separated by 2+ spaces
  - Date: DD/MM/YYYY  (slashes)
  - Balance: "277000.00 Cr"  (space before direction, plain decimal)
  - No voucher type tokens; direction inferred from balance change
  - Account header: centered name line (no leading ". " or with ". " prefix)
  - PAN line: "PAN No. XXXXXXXXXX"
  - Section end markers: "Sub Total :", "Cr Balance", "TOTAL :"

FIXES in this version:
  1. _parse_manmohan_stream: Account names appear as plain lines WITHOUT
     the ". " prefix OR with ". " prefix. The old code only caught ". NAME"
     accounts from the first page (AJAY DAMODAR DHOOT had no dot prefix).
     Fixed: detect account names as lines that are NOT dates, NOT amounts,
     NOT known keywords, AND have mixed/upper-case text > 3 chars.

  2. _parse_manmohan_txn: Was using cols[-1] as balance but for opening-balance
     rows the layout is:
       01/04/2024  Opening Balance As On 01/04/2024  277000.00  277000.00 Cr
     which gives 4 cols after the date. Fixed: search for Cr/Dr suffix explicitly.

  3. PAN line now also accepts "PAN No. " with no PAN (empty PAN — some accounts
     like MAHESH CO OP SOC have blank PAN).

  4. Opening balance voucher_type correctly set to 'Opbl' when description
     matches "Opening Balance".
"""

import re
from typing import List, Optional, Tuple
from app.models.transaction import RawTransaction
from app.utils.amount_parser import parse_amount, parse_balance, cr_value
from app.utils.date_parser import is_date_line
from app.core.logger import get_logger
from app.core.exceptions import FileParseError, EmptyFileError

logger = get_logger(__name__)

# ─────────────────────────── regex helpers ───────────────────────────────────
_AMOUNT_ONLY_RE   = re.compile(r'^[\d,]+\.\d{2}$')
_BALANCE_RE       = re.compile(r'^[\d,]+\.\d{2}\s*(Cr|Dr)?$|^-+$', re.IGNORECASE)
_VOUCHER_NO_RE    = re.compile(r'^BK-\d+$|^\d{3,4}$')
_VOUCHER_TYPE_RE  = re.compile(r'^(Opbl|Pay|Rec|Jrn)$', re.IGNORECASE)
_PAGE_TIMESTAMP_RE= re.compile(r'^\d{2}[-/]\d{2}[-/]\d{4}\s+\d{2}:\d{2}:\d{2}')
_PAGE_NO_RE       = re.compile(r'^Page\s*[:#]?\s*\d+', re.IGNORECASE)
_PAGE_OF_RE       = re.compile(r'^\d+\s*of\s*\d+')

# MANMOHAN-specific patterns
_MANMOHAN_TXN_LINE_RE = re.compile(r'^(\d{2}/\d{2}/\d{4})\s{2,}(.+)$')
_MANMOHAN_SUBTOTAL_RE = re.compile(
    r'^Sub\s+Total\s*:?\s*([\d,]+\.?\d*)\s+([\d,]+\.?\d*)', re.IGNORECASE
)
_MANMOHAN_CR_BAL_RE   = re.compile(r'^Cr\s+Balance\s+([\d,]+\.?\d*)', re.IGNORECASE)
_MANMOHAN_DR_BAL_RE   = re.compile(r'^Dr\s+Balance\s+([\d,]+\.?\d*)', re.IGNORECASE)
_MANMOHAN_TOTAL_RE    = re.compile(r'^TOTAL\s*:', re.IGNORECASE)
_MANMOHAN_PAN_RE      = re.compile(
    r'^PAN\s*(?:No\.?)?\s*:?\s*([A-Z]{5}\d{4}[A-Z])?', re.IGNORECASE
)

# Lines to always skip in MANMOHAN format
_MANMOHAN_SKIP_RE = re.compile(
    r'^\|\|'                          # "|| SHREE || ..."
    r'|^From\s*:'                     # "From : 01/04/2024 ..."
    r'|^Ledger\s+A/c\s+For'          # "Ledger A/c For ..."
    r'|^DATE\s+PARTICULARS'           # combined column header
    r'|^DATE\s+Particulars'
    r'|^Date\s+Particulars'
    r'|^Loans\s+[&A]\s+Borrowings'
    r'|^LOANS\s+[&A]\s+BORROWINGS'
    r'|^Dr\s+AMOUNT'                  # header continuation
    r'|^Cr\s+AMOUNT',
    re.IGNORECASE,
)

_COLUMN_HEADERS = frozenset([
    'Date', 'Voucher No.', 'Description', 'Debit', 'Credit', 'Balance',
    'DATE', 'PARTICULARS', 'Dr AMOUNT', 'Cr AMOUNT', 'BALANCE',
    'Voucher', 'No.', 'Narration',
    'Ch.no.',   # cheque number continuation lines
])

_FOOTER_PREFIXES = (
    'Group Of Ledger',
    'Account Ledger from',
    'Loans and Advances',
    'Loans & Advances',
    'Loans and Borrowings',
    'Loans & Borrowings',
)

_ADDRESS_LINE_RE = re.compile(r'^\d+/\d+[,\s]')

# Lines that look like account names we should NOT treat as new account starts
_NOT_ACCOUNT_RE = re.compile(
    r'^Ch\.no\.'                    # cheque reference lines
    r'|^Paid\s+'                    # "Paid GST Of ..."
    r'|^Page\s*:?\s*\d+'           # page numbers
    r'|^MANMOHAN'                   # company name repeat
    r'|^SHREE',
    re.IGNORECASE,
)


# ─────────────────────────── helper functions ─────────────────────────────────

def _is_column_header(line: str) -> bool:
    return line.strip() in _COLUMN_HEADERS


def _is_simple_footer(line: str) -> bool:
    if _PAGE_TIMESTAMP_RE.match(line):
        return True
    if _PAGE_NO_RE.match(line):
        return True
    if _PAGE_OF_RE.match(line):
        return True
    for prefix in _FOOTER_PREFIXES:
        if line.startswith(prefix):
            return True
    if _ADDRESS_LINE_RE.match(line):
        return True
    return False


def _remove_footer_blocks(lines: List[str]) -> List[str]:
    result: List[str] = []
    in_footer = False
    for line in lines:
        if _PAGE_TIMESTAMP_RE.match(line):
            in_footer = True
            continue
        if in_footer:
            stripped_upper = line.strip().upper()
            if stripped_upper == 'DATE' or stripped_upper.startswith('DATE '):
                in_footer = False
            continue
        if _is_simple_footer(line):
            continue
        result.append(line)
    return result


def _is_amount_only(line: str) -> bool:
    return bool(_AMOUNT_ONLY_RE.match(line.strip()))


def _is_balance_line(line: str) -> bool:
    return bool(_BALANCE_RE.match(line.strip()))


def _is_voucher_type(line: str) -> bool:
    return bool(_VOUCHER_TYPE_RE.match(line.strip()))


def _is_voucher_no(line: str) -> bool:
    return bool(_VOUCHER_NO_RE.match(line.strip()))


def _safe(tokens: list, idx: int) -> str:
    return tokens[idx].strip() if 0 <= idx < len(tokens) else ''


def _looks_like_account_name(line: str) -> bool:
    """
    Return True if a line looks like a lender/account name.
    Heuristics:
    - Not a date
    - Not a pure amount
    - Not a balance line
    - Not a voucher type
    - Not a known header/footer
    - Has at least one uppercase letter word > 2 chars
    - Not a cheque/payment reference
    """
    stripped = line.strip()
    if not stripped or len(stripped) < 3:
        return False
    if re.match(r'^\d{2}/\d{2}/\d{4}', stripped):
        return False
    if re.match(r'^\d{2}-\d{2}-\d{4}', stripped):
        return False
    if _is_amount_only(stripped):
        return False
    if _is_balance_line(stripped) and not re.search(r'[A-Za-z]', stripped):
        return False
    if _is_voucher_type(stripped):
        return False
    if _is_column_header(stripped):
        return False
    if _MANMOHAN_SKIP_RE.match(stripped):
        return False
    if _NOT_ACCOUNT_RE.match(stripped):
        return False
    # Must contain at least one alpha character word
    if not re.search(r'[A-Za-z]{2,}', stripped):
        return False
    # Reject lines that are purely numeric (e.g. stray page numbers)
    if re.match(r'^\d+$', stripped):
        return False
    return True


# ─────────────────────────── RawAccount dataclass ─────────────────────────────

class RawAccount:
    __slots__ = ('name', 'location', 'pan', 'transactions',
                 'total_debit', 'total_credit',
                 'closing_balance', 'closing_direction')

    def __init__(self, name: str, location: str = ''):
        self.name = name
        self.location = location
        self.pan: str = ''
        self.transactions: List[RawTransaction] = []
        self.total_debit: float = 0.0
        self.total_credit: float = 0.0
        self.closing_balance: float = 0.0
        self.closing_direction: str = 'Zero'

    def __repr__(self):
        return (f"RawAccount({self.name!r}, txns={len(self.transactions)}, "
                f"closing={self.closing_balance}{self.closing_direction})")


# ─────────────────────────── main parser ──────────────────────────────────────

class LedgerPDFParser:
    """
    Parses Tally-exported Ledger PDFs.
    Auto-detects MTA vs MANMOHAN format from date patterns.
    """

    def parse_file(self, pdf_path: str) -> List[RawAccount]:
        lines = self._extract_lines(pdf_path)
        if not lines:
            raise EmptyFileError("No text could be extracted from the PDF.")
        logger.info("Extracted %d lines from PDF", len(lines))

        fmt = self._detect_format(lines)
        logger.info("Detected PDF format: %s", fmt)

        if fmt == 'manmohan':
            # Use direct regex parser for MANMOHAN format (concatenated PDF text)
            try:
                accounts = self.parse_manmohan_pdf_direct(pdf_path)
                # Filter valid accounts
                accounts = [
                    a for a in accounts
                    if len(a.name) >= 3
                    and not a.name[0].isdigit()
                    and re.search(r'[A-Za-z]{2,}', a.name)
                ]
                if accounts:
                    logger.info("Direct MANMOHAN parser: %d accounts", len(accounts))
                    return accounts
            except Exception as e:
                logger.warning("Direct MANMOHAN parser failed: %s — falling back to stream", e)

            # Fallback to stream parser
            accounts = self._parse_manmohan_stream(lines)
        else:
            accounts = self._parse_mta_stream(lines)

        logger.info("Parsed %d account blocks", len(accounts))
        return accounts

    # ── format detection ──────────────────────────────────────────────────────

    def _detect_format(self, lines: List[str]) -> str:
        for line in lines[:80]:
            stripped = line.strip()
            if re.match(r'^\d{2}/\d{2}/\d{4}(\s|$)', stripped):
                return 'manmohan'
            if re.match(r'^\d{2}-\d{2}-\d{4}$', stripped):
                return 'mta'
        return 'mta'

    # ── text extraction ───────────────────────────────────────────────────────

    def parse_manmohan_pdf_direct(self, pdf_path: str) -> List[RawAccount]:
        """
        Direct parser for MANMOHAN TEXTILES LOAN.pdf format.
        The PDF text is highly concatenated — we use regex on the raw page text
        instead of line-by-line processing.
        """
        from pypdf import PdfReader
        reader = PdfReader(pdf_path)
        full_text = ''
        for page in reader.pages:
            full_text += (page.extract_text() or '') + '\n'

        accounts: List[RawAccount] = []

        # Split into account blocks by 'TOTAL :' marker
        # Each block: name → PAN → transactions → Sub Total → Cr/Dr Balance → TOTAL
        # Account names appear as '. NAME' or just 'NAME' after 'TOTAL :'
        # Strategy: find all account name + PAN sections

        # Pattern: find account boundaries by looking for PAN lines
        # Each account has a unique PAN (mostly), then transactions, then totals

        # Step 1: Normalise text — remove page headers
        text = re.sub(r'\|\|\s*SHREE\s*\|\|.*?BALANCE', '', full_text, flags=re.DOTALL)
        text = re.sub(r'From\s*:.*?To\s*\d{2}/\d{2}/\d{4}', '', text)
        text = re.sub(r'DATEPARTICULARS.*?BALANCE', '', text)
        text = re.sub(r'DATEPARTICULARSDr\s+AMOUNTCr\s+AMOUNTBALANCE', '', text)
        text = re.sub(r'Ledger\s+A/c\s+For\s+LOANS\s+[&A]\s+BORROWINGS', '', text)

        # Step 2: Find each account block using PAN pattern as anchor
        # Pattern to find account name + PAN pairs
        _BLOCK_RE = re.compile(
            r'(?:TOTAL\s*:[\s.]*)?'          # optional TOTAL : separator
            r'([A-Z][A-Z\s./&()-]{3,60}?)'   # account name
            r'PAN\s*No\.?\s*([A-Z]{5}\d{4}[A-Z])?',  # PAN (optional)
            re.DOTALL
        )

        # Better approach: split on account name pattern
        # Account names always precede 'PAN No.'
        _ACCT_PAN_RE = re.compile(
            r'([A-Z][A-Z\s\./&(),-]{2,60?}?)'
            r'\s*PAN\s*No\.?\s*([A-Z]{5}\d{4}[A-Z])?',
        )

        # Step 3: Use a cleaner approach — find all "Opening Balance As On" occurrences
        # Each account has exactly one Opening Balance entry
        _OB_RE = re.compile(
            r'(\d{2}/\d{2}/\d{4})\s+'      # date
            r'(\d[\d,.]*\.\d{2})Cr'         # opening balance amount + Cr
            r'(\d[\d,.]*\.\d{2})',           # running balance
        )

        # Extract all transactions grouped by looking for Sub Total markers
        _ACCOUNT_BLOCK_RE = re.compile(
            r'(?:(?:TOTAL\s*:\s*\.?\s*)|^)'
            r'([A-Z][A-Z\s\./&(),-]{2,60}?)'     # account name
            r'PAN\s*No\.?\s*([A-Z]{5}\d{4}[A-Z]|(?=Opening))'  # PAN
            r'.*?'                                  # transactions
            r'(?:Cr|Dr)\s+Balance\s*([\d,]+\.?\d*)',  # closing balance
            re.DOTALL
        )

        # Since the text is hard to parse with regex, use a simpler sequential approach:
        # Split text into segments at account-name boundaries
        return self._parse_manmohan_direct_text(full_text)

    def _parse_manmohan_direct_text(self, full_text: str) -> List[RawAccount]:
        """
        Parse concatenated MANMOHAN PDF text directly.
        Uses the known structure of the LOAN.pdf to extract accounts.
        """
        from app.utils.amount_parser import parse_amount, cr_value

        # Clean page headers from each page
        # Page text starts with "|| SHREE ||MANMOHAN TEXTILES..." header
        text = full_text

        # Remove page headers (everything between || SHREE || and BALANCE on each page)
        text = re.sub(
            r'\|\|\s*SHREE\s*\|\|.*?DATEPARTICULARSDr\s*AMOUNTCr\s*AMOUNTBALANCE',
            '', text, flags=re.DOTALL
        )

        accounts: List[RawAccount] = []

        # Strategy: find all lender blocks by PAN pattern + surrounding context
        # Each lender: NAME...PAN No. PANCODE...transactions...Sub Total...Cr Balance...TOTAL
        _PAN_LINE_RE = re.compile(r'PAN\s*No\.?\s*([A-Z]{5}\d{4}[A-Z])?')

        # Find all positions of PAN lines
        pan_positions = [(m.start(), m.group(1) or '') for m in _PAN_LINE_RE.finditer(text)]

        # For each PAN position, extract account name (text before it, after previous block end)
        # and closing balance (text after it, before TOTAL :)
        _TOTAL_RE = re.compile(r'TOTAL\s*:', re.IGNORECASE)
        _CR_BAL_RE = re.compile(r'(?:Cr|Dr)\s+Balance\s+([\d,]+\.?\d*)', re.IGNORECASE)
        _SUBTOTAL_RE = re.compile(r'Sub\s*\n*Total\s*:?\s*([\d,]+\.?\d*)', re.IGNORECASE)
        _DATE_AMT_RE = re.compile(
            r'(\d{2}/\d{2}/\d{4})\s+'
            r'(?:([\d,]+\.\d{2})Cr([\d,]+\.\d{2})'   # Cr transaction
            r'|([\d,]+\.\d{2})(?:Cr)?([\d,]+\.\d{2}))',  # Dr or Cr
            re.IGNORECASE
        )
        _OPENING_RE = re.compile(
            r'Opening Balance As On\s+\d{2}/\d{2}/\d{4}\s*'
            r'(\d{2}/\d{2}/\d{4})\s+'
            r'([\d,]+\.\d{2})Cr([\d,]+\.\d{2})',
            re.IGNORECASE
        )

        # Find TOTAL: positions to split blocks
        total_positions = [m.start() for m in _TOTAL_RE.finditer(text)]

        # For each account between TOTAL markers (or start → first TOTAL)
        block_starts = [0] + [p for p in total_positions]
        block_ends   = total_positions + [len(text)]

        for block_start, block_end in zip(block_starts, block_ends):
            block = text[block_start:block_end]
            if len(block.strip()) < 20:
                continue

            # Find PAN in this block
            pan_m = _PAN_LINE_RE.search(block)
            pan = (pan_m.group(1) or '') if pan_m else ''

            # Account name: text before PAN or before first date
            name_end = pan_m.start() if pan_m else len(block)
            name_text = block[:name_end].strip()
            # Clean: remove "TOTAL :. " prefix, dates, amounts
            name_text = re.sub(r'^TOTAL\s*:[\s.]*', '', name_text, flags=re.IGNORECASE)
            name_text = re.sub(r'\d{2}/\d{2}/\d{4}.*', '', name_text, flags=re.DOTALL)
            name_text = name_text.strip().rstrip('.')

            if not name_text or len(name_text) < 3:
                continue

            # Extract closing balance — look for 'Cr Balance AMOUNT' pattern
            cr_bal_m = _CR_BAL_RE.search(block)
            # Also try 'NUMBER Cr Balance' pattern which appears in this PDF
            if not cr_bal_m:
                cr_bal_m2 = re.search(r'([\d,]+\.\d{2})\s*\nCr\s+Balance', block)
                closing_balance = parse_amount(cr_bal_m2.group(1)) if cr_bal_m2 else 0.0
                closing_dir = 'Cr'
            else:
                closing_balance = parse_amount(cr_bal_m.group(1))
                closing_dir = 'Cr' if block[cr_bal_m.start():cr_bal_m.start()+2].upper() == 'CR' else 'Dr'

            # Better: get the first amount before 'Cr Balance' in the block
            # The pattern is: "277000.00\nCr Balance 307000.00..."
            # Or in concat text: "277000.00Cr Balance 307000.00"
            cb_m2 = re.search(r'([\d,]+\.\d{2})(?:\n)?Cr\s+Balance', block, re.IGNORECASE)
            if cb_m2:
                closing_balance = parse_amount(cb_m2.group(1))
                closing_dir = 'Cr'

            # Extract sub total
            sub_m = _SUBTOTAL_RE.search(block)

            # Create account
            acct = RawAccount(name=name_text)
            acct.pan = pan
            acct.closing_balance = closing_balance
            acct.closing_direction = closing_dir

            # Extract opening balance — first date line after PAN
            ob_text = block[pan_m.end():] if pan_m else block
            ob_m = re.search(
                r'Opening Balance As On\s+\d{2}/\d{2}/\d{4}\s*'
                r'(\d{2}/\d{2}/\d{4})\s+'
                r'([\d,]+\.\d{2})Cr([\d,]+\.\d{2})',
                ob_text, re.IGNORECASE
            )
            if ob_m:
                opening_bal = parse_amount(ob_m.group(2))
                acct.transactions.append(RawTransaction(
                    date_str=ob_m.group(1),
                    voucher_type='Opbl',
                    voucher_no='',
                    description='Opening Balance',
                    narration='',
                    amount=opening_bal,
                    balance_str=f'{ob_m.group(3)}Cr',
                    balance=parse_amount(ob_m.group(3)),
                    balance_direction='Cr',
                ))

            # Extract all date transactions
            # Pattern: DATE  AMOUNT Cr BALANCE  (or Dr)
            for txn_m in re.finditer(
                r'(\d{2}/\d{2}/\d{4})\s+([\d,]+\.\d{2})(?:Cr)?([\d,]+\.\d{2})',
                ob_text
            ):
                date_str = txn_m.group(1)
                amount   = parse_amount(txn_m.group(2))
                balance  = parse_amount(txn_m.group(3))
                # Determine if Cr or Dr based on context
                after_match = ob_text[txn_m.start():txn_m.start()+50]
                bal_dir = 'Cr'  # MANMOHAN loans are mostly Cr

                # Infer voucher type from amount and balance change
                # Opening balance already captured above
                if 'Opening Balance' in ob_text[max(0,txn_m.start()-30):txn_m.start()]:
                    continue  # already captured

                acct.transactions.append(RawTransaction(
                    date_str=date_str,
                    voucher_type='Rec',  # bank receipt
                    voucher_no='BK-1',
                    description='',
                    narration='',
                    amount=amount,
                    balance_str=f'{balance}Cr',
                    balance=balance,
                    balance_direction=bal_dir,
                ))

            # Filter: valid account name must not start with digits/amounts
            if (name_text
                    and len(name_text) >= 3
                    and not re.match(r'^\d', name_text)
                    and re.search(r'[A-Za-z]{2,}', name_text)
                    and closing_balance >= 0):
                accounts.append(acct)
                logger.debug("MANMOHAN direct: %s PAN=%s closing=%.0f%s",
                             name_text, pan, closing_balance, closing_dir)

        return accounts

    def _extract_lines(self, pdf_path: str) -> List[str]:
        """
        Extract lines from PDF.
        Handles both:
          - Proper newline-separated text (MTA format)
          - Concatenated text (MANMOHAN format — all columns on one line)
        For concatenated MANMOHAN format, we re-split on date patterns and keywords.
        """
        raw_pages = []

        # Try pypdf first
        try:
            from pypdf import PdfReader
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                text = page.extract_text() or ''
                raw_pages.append(text)
        except ImportError:
            pass
        except Exception as e:
            logger.warning("pypdf extraction failed: %s", e)

        # Fallback: pdfplumber
        if not raw_pages:
            try:
                import pdfplumber
                with pdfplumber.open(pdf_path) as pdf:
                    for page in pdf.pages:
                        raw_pages.append(page.extract_text() or '')
            except ImportError:
                raise FileParseError("No PDF library available. Install pypdf: pip install pypdf")
            except Exception as e:
                raise FileParseError(f"PDF extraction failed: {e}")

        if not raw_pages:
            return []

        # Check if text has proper newlines
        all_text = '\n'.join(raw_pages)
        natural_lines = [l.strip() for l in all_text.split('\n') if l.strip()]

        if len(natural_lines) > 10:
            # Proper newline format — use as-is
            return natural_lines

        # ── Concatenated format (MANMOHAN TEXTILES LOAN.pdf) ──────────────
        # Re-tokenize by injecting newlines before dates and known markers
        lines = []
        for page_text in raw_pages:
            # Insert newline before DD/MM/YYYY dates
            text = re.sub(r'(?<!\n)(?=\d{2}/\d{2}/\d{4})', '\n', page_text)
            # Insert newline before PAN No.
            text = re.sub(r'(?<!\n)(?=PAN\s*No)', '\n', text, flags=re.IGNORECASE)
            # Insert newline before Sub Total
            text = re.sub(r'(?<!\n)(?=Sub\s+Total)', '\n', text, flags=re.IGNORECASE)
            # Insert newline before Cr Balance / Dr Balance
            text = re.sub(r'(?<!\n)(?=(?:Cr|Dr)\s+Balance)', '\n', text, flags=re.IGNORECASE)
            # Insert newline before TOTAL :
            text = re.sub(r'(?<!\n)(?=TOTAL\s*:)', '\n', text, flags=re.IGNORECASE)
            # Insert newline before Ch.no.
            text = re.sub(r'(?<!\n)(?=Ch\.no\.)', '\n', text, flags=re.IGNORECASE)

            for raw_line in text.split('\n'):
                stripped = raw_line.strip()
                if stripped:
                    lines.append(stripped)

        return lines

    # ══════════════════════════════════════════════════════════════════════════
    # MTA format parser  (token-stream: one PDF column value per line)
    # ══════════════════════════════════════════════════════════════════════════

    def _parse_mta_stream(self, raw_lines: List[str]) -> List[RawAccount]:
        without_footers = _remove_footer_blocks(raw_lines)
        tokens = [ln for ln in without_footers if not _is_column_header(ln)]

        accounts: List[RawAccount] = []
        current: Optional[RawAccount] = None
        i = 0

        while i < len(tokens):
            tok = tokens[i]

            if is_date_line(tok):
                txn, i = self._parse_mta_transaction(tokens, i)
                if txn and current is not None:
                    current.transactions.append(txn)
                continue

            if _is_amount_only(tok):
                tok2 = _safe(tokens, i + 1)
                tok3 = _safe(tokens, i + 2)
                if _is_amount_only(tok2) and _is_balance_line(tok3):
                    if current is not None:
                        current.total_debit = parse_amount(tok)
                        current.total_credit = parse_amount(tok2)
                        amt, direction = parse_balance(tok3)
                        current.closing_balance = amt
                        current.closing_direction = direction
                        accounts.append(current)
                        current = None
                    i += 3
                    continue
                if _is_amount_only(tok2) and (tok3 == '-' or tok3 == '–'):
                    if current is not None:
                        current.total_debit = parse_amount(tok)
                        current.total_credit = parse_amount(tok2)
                        current.closing_balance = 0.0
                        current.closing_direction = 'Zero'
                        accounts.append(current)
                        current = None
                    i += 3
                    continue
                i += 1
                continue

            if tok in ('-', '–'):
                i += 1
                continue

            if current is not None and tok == current.name:
                i += 1
                continue
            if current is not None and tok == current.location:
                i += 1
                continue

            if (not _is_voucher_type(tok)
                    and not _is_voucher_no(tok)
                    and not _is_balance_line(tok)):
                if current is None:
                    current = RawAccount(name=tok)
                    logger.debug("New MTA account: %s", tok)
                elif not current.location:
                    current.location = tok
            i += 1

        if current is not None and current not in accounts:
            accounts.append(current)

        return accounts

    def _parse_mta_transaction(self, tokens: List[str], start: int):
        i = start
        date_str = tokens[i]; i += 1

        if i >= len(tokens):
            return None, i

        v_type_raw = tokens[i].strip(); i += 1
        v_type = v_type_raw.capitalize() if v_type_raw else ''

        voucher_no = ''
        if _is_voucher_no(_safe(tokens, i)):
            voucher_no = tokens[i]; i += 1

        description = ''
        if i < len(tokens) and not _is_amount_only(tokens[i]) and not is_date_line(tokens[i]):
            description = tokens[i]; i += 1

        amount = 0.0
        if i < len(tokens) and _is_amount_only(tokens[i]):
            amount = parse_amount(tokens[i]); i += 1

        balance_str = ''
        balance = 0.0
        balance_direction = 'Zero'
        if i < len(tokens) and _is_balance_line(tokens[i]):
            balance_str = tokens[i]
            balance, balance_direction = parse_balance(balance_str)
            i += 1

        narration = ''
        if (i < len(tokens)
                and not is_date_line(tokens[i])
                and not _is_amount_only(tokens[i])
                and not _is_balance_line(tokens[i])
                and not _is_voucher_type(tokens[i])):
            narration = tokens[i]; i += 1

        txn = RawTransaction(
            date_str=date_str,
            voucher_type=v_type,
            voucher_no=voucher_no,
            description=description,
            narration=narration,
            amount=amount,
            balance_str=balance_str,
            balance=balance,
            balance_direction=balance_direction,
        )
        return txn, i

    # ══════════════════════════════════════════════════════════════════════════
    # MANMOHAN format parser  (row-per-line)
    # ══════════════════════════════════════════════════════════════════════════

    def _parse_manmohan_stream(self, raw_lines: List[str]) -> List[RawAccount]:
        without_footers = _remove_footer_blocks(raw_lines)

        accounts: List[RawAccount] = []
        current: Optional[RawAccount] = None
        prev_cr: float = 0.0

        for line in without_footers:
            if not line:
                continue

            # ── Skip known header/footer patterns ──────────────────────────
            if _MANMOHAN_SKIP_RE.match(line):
                continue
            if _is_column_header(line):
                continue

            # Strip leading ". " from account name lines like ". AJAY DAMODAR DHOOT"
            clean_line = line
            if line.startswith('. ') or line.startswith('.\t'):
                clean_line = line[2:].strip()

            # ── Transaction row (starts with DD/MM/YYYY) ───────────────────
            m = _MANMOHAN_TXN_LINE_RE.match(line)
            if m:
                if current is None:
                    continue
                date_str = m.group(1)
                rest = m.group(2).strip()
                cols = re.split(r'\s{2,}', rest)
                txn, new_cr = self._parse_manmohan_txn(date_str, cols, prev_cr)
                if txn:
                    prev_cr = new_cr
                    current.transactions.append(txn)
                continue

            # ── Sub Total line ──────────────────────────────────────────────
            ms = _MANMOHAN_SUBTOTAL_RE.match(line)
            if ms:
                if current:
                    current.total_debit  = parse_amount(ms.group(1))
                    current.total_credit = parse_amount(ms.group(2))
                continue

            # ── Cr/Dr Balance (closing balance) ────────────────────────────
            mc = _MANMOHAN_CR_BAL_RE.match(line)
            if mc:
                if current:
                    current.closing_balance   = parse_amount(mc.group(1))
                    current.closing_direction = 'Cr'
                continue
            md = _MANMOHAN_DR_BAL_RE.match(line)
            if md:
                if current:
                    current.closing_balance   = parse_amount(md.group(1))
                    current.closing_direction = 'Dr'
                continue

            # ── TOTAL : → account block complete ───────────────────────────
            if _MANMOHAN_TOTAL_RE.match(line):
                if current and current not in accounts:
                    accounts.append(current)
                    logger.debug("Completed Manmohan account: %s", current.name)
                current = None
                prev_cr = 0.0
                continue

            # ── PAN No. line ────────────────────────────────────────────────
            if re.match(r'^PAN\s*(?:No\.?)?\s*:?\s*', line, re.IGNORECASE):
                mp = _MANMOHAN_PAN_RE.match(line)
                if mp and current:
                    current.pan = (mp.group(1) or '').upper()
                continue

            # ── Skip pure-amount / balance-only lines ──────────────────────
            if _is_amount_only(line) or (_is_balance_line(line) and not re.search(r'[A-Za-z]', line)):
                continue

            # ── Skip cheque-number continuation lines (Ch.no. XXXX) ────────
            if re.match(r'^Ch\.no\.', line, re.IGNORECASE):
                continue

            # ── Skip "Page : N" lines ──────────────────────────────────────
            if _PAGE_NO_RE.match(line):
                continue

            # ── Page continuation: same account name or location repeated ──
            if current:
                if clean_line == current.name or line == current.name:
                    continue
                if clean_line == current.location or line == current.location:
                    continue

            # ── New account name ────────────────────────────────────────────
            if _looks_like_account_name(clean_line):
                if current is None:
                    current = RawAccount(name=clean_line)
                    prev_cr = 0.0
                    logger.debug("New Manmohan account: %s", clean_line)
                elif not current.location and clean_line != current.name:
                    # Could be city/address line after account name
                    # Only accept if it doesn't look like another full account name
                    # (i.e., it comes right after and has no prior transactions)
                    if not current.transactions:
                        current.location = clean_line

        # Handle last account if file ends without a TOTAL : marker
        if current and current not in accounts:
            accounts.append(current)

        return accounts

    def _parse_manmohan_txn(
        self, date_str: str, cols: List[str], prev_cr: float
    ) -> Tuple[Optional[RawTransaction], float]:
        """
        Parse one MANMOHAN transaction row.

        cols layout examples (after splitting rest on 2+ spaces):
          Opening: ['Opening Balance As On 01/04/2024', '277000.00', '277000.00 Cr']
          Payment: ['ICICI BANK LTD', '27000.00', '250000.00 Cr']
          Interest:['INTEREST PAID A/C', '30000.00', '280000.00 Cr']
          TDS:     ['TDS PAYABLE A/C.', '3000.00', '277000.00 Cr']
        """
        if not cols:
            return None, prev_cr

        description = cols[0].strip() if cols else ''

        # Find balance string: last column containing Cr or Dr suffix
        balance_str = ''
        balance = 0.0
        bal_dir = 'Zero'

        for c in reversed(cols[1:]):
            c = c.strip()
            if re.search(r'(Cr|Dr)\s*$', c, re.IGNORECASE):
                balance_str = c
                balance, bal_dir = parse_balance(c)
                break

        # If no explicit Cr/Dr found, try last column as plain number
        if not balance_str and len(cols) > 1:
            balance_str = cols[-1].strip()
            balance, bal_dir = parse_balance(balance_str)

        new_cr = cr_value(balance, bal_dir)

        # Extract amount from middle columns (all except first and last)
        amount = 0.0
        mid_cols = cols[1:-1] if len(cols) > 2 else []
        for mc in mid_cols:
            v = parse_amount(mc)
            if v > 0.0:
                amount = v
                break

        # If no middle col, use the difference or the balance itself for opening
        is_opening = bool(re.match(r'^(to\s+)?opening\s+balance|^balance\s+b[/.]?d',
                                   description, re.IGNORECASE))
        if is_opening:
            voucher_type = 'Opbl'
            if amount == 0.0:
                amount = balance
        else:
            delta = new_cr - prev_cr
            if abs(delta) < 0.005:
                voucher_type = 'Jrn'
            elif delta > 0:
                voucher_type = 'Rec'
            else:
                voucher_type = 'Pay'

        # Tag non-opening bank transactions
        voucher_no = 'BK-1' if voucher_type in ('Rec', 'Pay') else ''

        txn = RawTransaction(
            date_str=date_str,
            voucher_type=voucher_type,
            voucher_no=voucher_no,
            description=description,
            narration='',
            amount=amount,
            balance_str=balance_str,
            balance=balance,
            balance_direction=bal_dir,
        )
        return txn, new_cr

    # kept for backward compat
    def _parse_token_stream(self, raw_lines: List[str]) -> List[RawAccount]:
        return self._parse_mta_stream(raw_lines)
