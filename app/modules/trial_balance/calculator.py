"""
Trial Balance Calculator with intelligent group normalisation.
Maps raw parsed group names → canonical Genius Software group names.
"""
import os
import json
from typing import List, Dict, Optional
from dataclasses import dataclass, field
from app.modules.trial_balance.parser import TrialBalanceEntry
from app.core.logger import get_logger

logger = get_logger(__name__)

# ── Master reference grouping dictionary ──────────────────────────────────────
# Built from the client's "Master Reference For Grouping" workbook: an
# authoritative Ledger Name → Group map plus the canonical Group vocabulary.
# Context-dependent names (Cloth/Yarn/Discount, which vary by Trading section)
# are deliberately excluded and resolved from the source section header instead.
_MASTER_LEDGER_MAP: Dict[str, str] = {}
_MASTER_GROUPS: set = set()
try:
    _mpath = os.path.join(os.path.dirname(__file__), 'grouping_master.json')
    with open(_mpath, 'r', encoding='utf-8') as _fh:
        _mdata = json.load(_fh)
    _MASTER_LEDGER_MAP = {k.strip().lower(): v.strip().upper()
                          for k, v in _mdata.get('ledger_to_group', {}).items()}
    _MASTER_GROUPS = {g.strip().upper() for g in _mdata.get('groups', [])}
    logger.info("Loaded %d master ledger mappings, %d canonical groups",
                len(_MASTER_LEDGER_MAP), len(_MASTER_GROUPS))
except Exception as _e:  # pragma: no cover
    logger.warning("Master grouping reference not loaded: %s", _e)


@dataclass
class GroupSummary:
    group_name: str
    entries: List[TrialBalanceEntry] = field(default_factory=list)
    total_debit: float = 0.0
    total_credit: float = 0.0
    net_balance: float = 0.0


# ── Canonical group aliases (lower-case key → Genius group name) ───────────────
_ALIASES: Dict[str, str] = {
    # Capital
    'capital a/c':                    'CAPITAL',
    'capital account':                 'CAPITAL',
    'capital':                         'CAPITAL',
    'net profit':                      'CAPITAL',
    'net loss':                        'CAPITAL',
    'profit & loss a/c':               'CAPITAL',
    'profit and loss a/c':             'CAPITAL',
    'profit & loss account':           'CAPITAL',
    'p & l a/c':                       'CAPITAL',
    'shree ganesh ji maharaj':         'CAPITAL',
    # FIX: SHRI GANESHJI MAHARAJ (common Tally entry for proprietor share in profit)
    'shri ganeshji maharaj':           'CAPITAL',
    'ganeshji maharaj':                'CAPITAL',

    # Unsecured Loans
    'loans & borrowings':              'UNSECURED LOANS',
    'loan & borrowings':               'UNSECURED LOANS',
    'loans and borrowings':            'UNSECURED LOANS',
    'unsecured loans':                 'UNSECURED LOANS',
    'unsecured loan':                  'UNSECURED LOANS',
    'loan (liability)':                'UNSECURED LOANS',
    # FIX: HDFC HOME LOAN is a personal loan = Unsecured Loan per CA requirement
    'hdfc home loan':                  'UNSECURED LOANS',
    'hdfc home loan a/c':              'UNSECURED LOANS',

    # Secured Loans / OD
    'secured loans':                   'SECURED LOANS',
    'bank occ a/c':                    'SECURED LOANS',
    'bank od a/c':                     'SECURED LOANS',
    'bank occ':                        'SECURED LOANS',

    # Sundry Creditors (all sub-groups → SUNDRY CREDITORS)
    'sundry creditors':                'SUNDRY CREDITORS',
    'sundry creditors for yarn':       'SUNDRY CREDITORS',
    'sundry creditors for cloth':      'SUNDRY CREDITORS',
    'sundry creditors for sizing':     'SUNDRY CREDITORS',
    'sundry creditors for weaving':    'SUNDRY CREDITORS',
    'sundry creditors for expenses':   'SUNDRY CREDITORS',
    'sundry creditors (pur.)':         'SUNDRY CREDITORS',
    'sundry brockerage':               'SUNDRY CREDITORS',
    'broker a/c':                      'SUNDRY CREDITORS',
    'brokers':                         'SUNDRY CREDITORS',
    'karkhandar a/c.':                 'SUNDRY CREDITORS',
    'karkhandar auto':                 'SUNDRY CREDITORS',

    # Sundry Debtors
    'sundry debtors':                  'SUNDRY DEBTORS',
    'sundry debtors (sale)':           'SUNDRY DEBTORS',

    # Provisions / Duties & Taxes
    'provisions':                      'PROVISIONS',
    'duties and taxes':                'PROVISIONS',
    'duty and taxes':                  'DUTIES AND TAXES',
    'duty & taxes':                    'DUTIES AND TAXES',
    # Expenditure / Expense group names (common misspellings)
    'expinditure a/c':                 'INDIRECT EXPENSES',
    'expenditure a/c':                 'INDIRECT EXPENSES',
    'expinditure':                     'INDIRECT EXPENSES',
    'expenditure':                     'INDIRECT EXPENSES',
    # Net profit/loss entries → Capital
    'z. net profit a/c':               'CAPITAL',
    'net profit a/c':                  'CAPITAL',
    'net loss a/c':                    'CAPITAL',
    # Loan group names
    'sarafi given':                    'UNSECURED LOANS',
    'sarafi':                          'UNSECURED LOANS',
    # Debtor group names
    'debtors for cloth':               'SUNDRY DEBTORS',
    'debtors for yarn':                'SUNDRY DEBTORS',
    'debtors for sales':               'SUNDRY DEBTORS',
    # FIX: CURRENT LIABILITIES group in BS format maps to PROVISIONS (liability side)
    'current liabilities':             'PROVISIONS',
    'loans (liability)':               'UNSECURED LOANS',
    'loans(liability)':                'UNSECURED LOANS',
    'loans & borrowings':              'UNSECURED LOANS',
    'bank accounts':                   'CASH AND BANK',
    'bank account':                    'CASH AND BANK',
    'cash-in-hand':                    'CASH IN HAND',
    'sales accounts':                  'SALES A/C',
    'sales account':                   'SALES A/C',
    'purchase accounts':               'PURCHASE A/C',
    'purchase account':                'PURCHASE A/C',
    'direct expenses':                 'DIRECT EXPENSES (M)',
    'indirect incomes':                'INDIRECT INCOMES',
    'indirect expenses':               'INDIRECT EXPENSES',
    # Reserve & Surplus = secured bank loan (term loan) in Tally BS
    'reserve & surplus':               'SECURED LOANS',
    'reserves & surplus':              'SECURED LOANS',
    # Broker A/C = Sundry Creditors (brokerage payable)
    'broker a/c':                      'SUNDRY CREDITORS',
    'brokers a/c':                     'SUNDRY CREDITORS',
    # Labour & Advances = Loans and Advances (Assets)
    'labour & advances':               'LOANS AND ADVANCES (ASSETS)',
    'labour and advances':             'LOANS AND ADVANCES (ASSETS)',
    'labour advances':                 'LOANS AND ADVANCES (ASSETS)',

    # Fixed Assets
    'fixed assets':                    'FIXED ASSETS',
    'fixed asset':                     'FIXED ASSETS',

    # Investments
    'investments':                     'INVESTMENTS',
    'investment':                      'INVESTMENTS',

    # Shares — FIX: SHARES VSB A/C is a share investment = SHARES group
    'shares':                          'SHARES',
    'equity shares':                   'EQUITY SHARES',

    # Deposits
    'deposits':                        'DEPOSITS',

    # Loans & Advances (Assets)
    'loans & advances':                'LOANS AND ADVANCES (ASSETS)',
    'loans and advances':              'LOANS AND ADVANCES (ASSETS)',
    'loans and advances (assets)':     'LOANS AND ADVANCES (ASSETS)',
    'loans and advances (assets)':     'LOANS AND ADVANCES (ASSETS)',
    'loans and advances (asset)':      'LOANS AND ADVANCES (ASSETS)',
    'loan (asset)':                    'LOANS AND ADVANCES (ASSETS)',
    'advances':                        'LOANS AND ADVANCES (ASSETS)',

    # Cash & Bank — FIX: BANK ACCOUNT group (used in UGARARAM BS) → CASH AND BANK
    'cash and bank':                   'CASH AND BANK',
    'cash & bank':                     'CASH AND BANK',
    'bank a/c':                        'CASH AND BANK',
    'bank account':                    'CASH AND BANK',
    'bank':                            'CASH AND BANK',
    'balance at bank/cash':            'CASH AND BANK',
    'balance at bank / cash':          'CASH AND BANK',
    'cash & bank balance':             'CASH AND BANK',
    'cash and bank balance':           'CASH AND BANK',

    # Cash In Hand
    'cash in hand':                    'CASH IN HAND',

    # Other Current Assets — FIX: CURRENT ASSETS group (asset side) → OTHER CURRENT ASSETS
    'other current assets':            'OTHER CURRENT ASSETS',
    'current assets':                  'OTHER CURRENT ASSETS',
    'duties and taxes (asset)':        'OTHER CURRENT ASSETS',

    # Balance with Revenue Authority
    'balance with revenue authority':  'BALANCE WITH REVENUE AUTHORITY',
    'tds receivable':                  'BALANCE WITH REVENUE AUTHORITY',

    # Opening Stock
    'opening stock':                   'OPENING STOCK',
    'stock in hand ( opening )':       'OPENING STOCK',
    'stock in hand (opening)':         'OPENING STOCK',

    # Closing Stock — excluded by service layer per rules, kept for completeness
    'closing stock':                   'CLOSING STOCK',
    'stock in hand (closing)':         'CLOSING STOCK',
    'stock in hand ( closing )':       'CLOSING STOCK',

    # Sales — all common variations
    'sales a/c':                       'SALES A/C',
    'sale a/c':                        'SALES A/C',
    'sale account':                    'SALES A/C',
    'sales':                           'SALES A/C',
    'cloth sale a/c':                  'SALES A/C',
    'cloth sales a/c':                 'SALES A/C',
    'cloth sale':                      'SALES A/C',
    'yarn sale a/c':                   'SALES A/C',
    'yarn sales a/c':                  'SALES A/C',
    'yarn sale':                       'SALES A/C',
    'sample cloth sale a/c':           'SALES A/C',
    'yarn sale ( sized yarn )':        'SALES A/C',
    # Closing stock (right/income side of P&L)
    'stock in hand':                   'CLOSING STOCK',
    'stock in hand a/c':               'CLOSING STOCK',
    # Direct & indirect income (right side of P&L)
    'cloth jobwork received':          'DIRECT INCOMES',
    'job weaving charges (claim)':     'DIRECT INCOMES',
    'sizing warping charges (claim)':  'DIRECT INCOMES',
    'direct expinditure':              'DIRECT INCOMES',  # right-side claim income
    'indirect expences':               'INDIRECT INCOMES',
    'indirect income':                 'INDIRECT INCOMES',
    'other income':                    'INDIRECT INCOMES',
    'gst round off':                   'INDIRECT INCOMES',
    'intrest late bill ( gst )':       'INDIRECT INCOMES',
    'interest on income tax':          'INDIRECT INCOMES',
    'rate difference (sale/purchase':  'INDIRECT INCOMES',
    'sundry creditors write-off':      'INDIRECT INCOMES',
    'discount a/c':                    'INDIRECT INCOMES',

    # Purchases
    'purchase a/c':                    'PURCHASE A/C',
    'purchase account':                'PURCHASE A/C',
    'purchases':                       'PURCHASE A/C',

    # Wages (trading concern) → Trading A/c direct expense
    'wages':                           'DIRECT EXPENSES (M)',
    'wages & salary':                  'DIRECT EXPENSES (M)',
    'wages and salary':                'DIRECT EXPENSES (M)',

    # Common indirect expenses seen without a Group column
    'electricity expenses':            'INDIRECT EXPENSES',
    'electricity expense':             'INDIRECT EXPENSES',
    'electricity':                     'INDIRECT EXPENSES',
    'electricity charges':             'INDIRECT EXPENSES',

    # Plant & Machinery → Fixed Asset
    'plant & machinery':               'FIXED ASSETS',
    'plant and machinery':             'FIXED ASSETS',
    'plant & machinery a/c':           'FIXED ASSETS',

    # Advance Tax → Balance with Revenue Authority (current asset)
    'advance tax':                     'BALANCE WITH REVENUE AUTHORITY',
    'advance income tax':              'BALANCE WITH REVENUE AUTHORITY',

    # Manufacturing
    'manufacturing expenses':          'MANUFACTURING EXPENSES',

    # ── Master-reference ledger mappings (specific groups override the
    #    generic expense-indicator guard) ───────────────────────────────────
    'brokerage':                       'COMMISSION PAID',
    'commission':                      'COMMISSION PAID',
    'bank charges':                    'COMMISSION PAID',
    'bank commission':                 'COMMISSION PAID',
    'salary':                          'COMPENSATION TO EMPLOYEES',
    'salary a/c':                      'COMPENSATION TO EMPLOYEES',
    'salary a/c.':                     'COMPENSATION TO EMPLOYEES',
    'bonus':                           'COMPENSATION TO EMPLOYEES',
    'allowances':                      'COMPENSATION TO EMPLOYEES',
    'audit fees':                      'INDIRECT EXPENSES',
    'gst payment':                     'DUTIES AND TAXES',
    'travelling exp.':                 'TRAVELLING',
    'travelling':                      'TRAVELLING',
    'transport exp.':                  'FREIGHT',
    'transport charges':               'FREIGHT',
    'hamali exp.':                     'FREIGHT',
    'hamali':                          'FREIGHT',
    'vahatuk':                         'FREIGHT',
    'carriage outward':                'FREIGHT',
    'interest paid a/c.':              'INTEREST PAID',
    'interest received a/c':           'INTEREST RECEIVED',
    'vehicle insurance':               'INSURANCE',
    'stock insuranse':                 'INSURANCE',
    'depreciation':                    'DEPRECIATION',
    'bad debts':                       'BAD DEBTS',
    'donation':                        'DONATION',
    # Direct (trading) expenses per master reference
    'checking charges':                'DIRECT EXPENSES (M)',
    'mending charges':                 'DIRECT EXPENSES (M)',
    'packing charges':                 'DIRECT EXPENSES (M)',
    'butta cutting charges':           'DIRECT EXPENSES (M)',
    'design pattern charges':          'DIRECT EXPENSES (M)',
    'folding':                         'DIRECT EXPENSES (M)',
    'winding':                         'DIRECT EXPENSES (M)',
    'wiInding expenses':               'DIRECT EXPENSES (M)',
    'widing expenses':                 'DIRECT EXPENSES (M)',
    'warping charges':                 'DIRECT EXPENSES (M)',
    'sizing warping charges':          'DIRECT EXPENSES (M)',

    # Direct Expenses (Manufacturing)
    'direct expenses (m)':             'DIRECT EXPENSES (M)',
    'direct expenses':                 'DIRECT EXPENSES (M)',
    'direct expinditure':              'DIRECT EXPENSES (M)',
    'direct expenditure':              'DIRECT EXPENSES (M)',

    # Direct Incomes
    'direct incomes':                  'DIRECT INCOMES',
    'direct income':                   'DIRECT INCOMES',

    # Indirect
    'indirect expenses':               'INDIRECT EXPENSES',
    'indirect income':                 'INDIRECT INCOMES',
    'indirect incomes':                'INDIRECT INCOMES',

    # Specific expense/income groups
    'commission paid':                 'COMMISSION PAID',
    'commission received':             'COMMISSION RECEIVED',
    'compensation to employees':       'COMPENSATION TO EMPLOYEES',
    'depreciation':                    'DEPRECIATION',
    'interest paid':                   'INTEREST PAID',
    'interest received':               'INTEREST RECEIVED',
    'freight':                         'FREIGHT',
    'carriage inward':                 'CARRIAGE INWARD',
    'insurance':                       'INSURANCE',
    'telephone':                       'TELEPHONE',
    'travelling':                      'TRAVELLING',
    'conveyance':                      'CONVEYANCE',
    'rent':                            'RENT',
    'rates and taxes':                 'RATES AND TAXES',
    'power and fuel':                  'POWER AND FUEL',
    'power and fuel (m)':              'POWER AND FUEL (M)',
    'advertisement':                   'ADVERTISEMENT',
    'bad debts':                       'BAD DEBTS',
    'donation':                        'DONATION',
    'other expenses':                  'OTHER EXPENSES',
    'other incomes':                   'OTHER INCOMES',
    'other income':                    'OTHER INCOMES',
    'auditors remuneration':           'AUDITORS REMUNERATION',
    'repair & maintenance':            'REPAIR & MAINTENANCE',
    'staff welfare':                   'STAFF WELFARE',
    'financial expenses':              'FINANCIAL EXPENSES',
    'extra-ordinary expenses':         'EXTRA-ORDINARY EXPENSES',
}

# Keyword-based fallback matching (substring checks)
_KEYWORD_MAP = [
    (['sarafi', 'family loan', 'personal loan', 'director loan', 'friend loan',
      'home loan'],
     'UNSECURED LOANS'),
    (['sundry creditor', 'karkhandar'],
     'SUNDRY CREDITORS'),
    (['sundry debtor'],
     'SUNDRY DEBTORS'),
    (['fixed asset'],
     'FIXED ASSETS'),
    (['opening stock'],
     'OPENING STOCK'),
    (['closing stock'],
     'CLOSING STOCK'),
    (['bank od', 'bank occ', 'bank oc'],
     'SECURED LOANS'),
    (['cash in hand'],
     'CASH IN HAND'),
    (['cash and bank', 'cash & bank'],
     'CASH AND BANK'),
    (['loans and advance', 'loans & advance', 'loans / advance'],
     'LOANS AND ADVANCES (ASSETS)'),
    (['manufacture', 'manufacturing'],
     'MANUFACTURING EXPENSES'),
    (['purchase'],
     'PURCHASE A/C'),
    (['sale'],
     'SALES A/C'),
    (['ganesh', 'ganeshji'],
     'CAPITAL'),
    # Common P&L expense patterns (last-resort, after exact aliases/master map).
    (['repair', 'maintenance', 'maintainance'],
     'REPAIR & MAINTENANCE'),
    (['interest'],
     'INTEREST PAID'),
    (['godown rent', 'office rent', 'rent &', 'rent and', 'rent a/c'],
     'RENT'),
    (['wages'],
     'DIRECT WAGES'),
    (['commission'],
     'COMMISSION PAID'),
    (['profession tax', 'professional tax'],
     'RATES AND TAXES'),
    (['donation'],
     'DONATION'),
    (['insurance'],
     'INSURANCE'),
    (['depreciation'],
     'DEPRECIATION'),
    (['freight', 'cartage', 'hamali', 'vahatook', 'vahatuk', 'transport'],
     'FREIGHT'),
]

# ── Context-aware group inference ─────────────────────────────────────────────
# When a ledger/account name matches none of the above, we try to infer group
# from the CONTEXT GROUP that the parser supplied. This handles cases where
# Format J correctly sets group="SUNDRY CREDITORS" for person names, but the
# name itself has no keyword signal. The _normalise function uses the raw group
# first (which IS the canonical group from the parser), so this is automatic.
# The mapping below additionally handles "As Per Schedule X" references whose
# group is the section header name.

_SCHEDULE_SECTION_MAP: Dict[str, str] = {
    # Map section-header group names → canonical Genius groups
    # These are set by Format J parser as raw group names
    'CAPITAL': 'CAPITAL',
    'CAPITAL ACCOUNT': 'CAPITAL',
    'PARTNERS CAPITAL': 'CAPITAL',
    'CURRENT LIABILITIES': 'PROVISIONS',
    'DUTIES AND TAXES': 'DUTIES AND TAXES',
    'BROKERS (SALE COMMISSION AGENTS)': 'SUNDRY CREDITORS',
    'PROVISION': 'PROVISIONS',
    'SUNDRY CREDITORS': 'SUNDRY CREDITORS',
    'SUNDRY CREDITORS MILL': 'SUNDRY CREDITORS',
    'SUNDRY PAYABLES': 'SUNDRY CREDITORS',
    'LOANS LIABILITIES': 'UNSECURED LOANS',
    'CURRENT ASSETS': 'OTHER CURRENT ASSETS',
    'CASH & BANK BALANCES': 'CASH AND BANK',
    'CASH AND BANK BALANCES': 'CASH AND BANK',
    'LOANS / ADVANCES A/C': 'LOANS AND ADVANCES (ASSETS)',
    'LOANS AND ADVANCES A/C': 'LOANS AND ADVANCES (ASSETS)',
    'SUNDRY DEBTORS': 'SUNDRY DEBTORS',
    'SUNDRY RECEIVABLES': 'SUNDRY DEBTORS',
    'FIXED ASSETS': 'FIXED ASSETS',
    'OTHER CURRENT ASSETS': 'OTHER CURRENT ASSETS',
    'STOCK IN HAND': 'CLOSING STOCK',
    'OPENING STOCK': 'OPENING STOCK',
    'CLOSING STOCK': 'CLOSING STOCK',
    'EXPENSES DIRECT': 'DIRECT EXPENSES (M)',
    'EXPENSES DIRECT (T&M)': 'DIRECT EXPENSES (M)',
    'EXPENSES INDIRECT': 'INDIRECT EXPENSES',
    'EXPENSES INDIRECT (P&L)': 'INDIRECT EXPENSES',
    'OTHER INCOME': 'INDIRECT INCOMES',
    'INDIRECT INCOME': 'INDIRECT INCOMES',
    'INDIRECT INCOMES': 'INDIRECT INCOMES',
    'SALES': 'SALES A/C',
    'SALES A/C': 'SALES A/C',
    'PURCHASE': 'PURCHASE A/C',
    'PURCHASE A/C': 'PURCHASE A/C',
    'DIRECT EXPENSES (M)': 'DIRECT EXPENSES (M)',
    'INDIRECT EXPENSES': 'INDIRECT EXPENSES',
    'INVESTMENTS': 'INVESTMENTS',
}



# Words that unambiguously mark a Profit & Loss EXPENSE ledger.
# These must win over asset/creditor keyword guesses (e.g. "Computer Expenses"
# is an expense, NOT a fixed asset; "Brokerage" is an expense, NOT a creditor).
_EXPENSE_INDICATORS = (
    'expenses', 'expense', 'exp.', ' exp', 'charges', 'charge',
    'fees', 'fee', ' bill', 'bonus',
)
# Standalone ASSET ledger names (exact, after expense guard).
_ASSET_TERMS = {
    'computer', 'computers', 'furniture', 'furniture & fixtures', 'machinery',
    'plant', 'plant & machinery', 'building', 'buildings', 'vehicle', 'vehicles',
    'equipment', 'equipments', 'land', 'office equipment',
}


def _match(text: str) -> Optional[str]:
    """Return the canonical group for a single string, or None if unknown."""
    if not text:
        return None
    key = text.strip().lower()
    # 1) Exact alias wins.
    if key in _ALIASES:
        return _ALIASES[key]
    # 1b) Master reference ledger→group map (client's authoritative dictionary).
    if key in _MASTER_LEDGER_MAP:
        return _MASTER_LEDGER_MAP[key]
    # 2) Balance-sheet PARTY/LIABILITY indicator outranks the expense guard:
    #    "Sundry Creditors for Expenses" (often PDF-truncated to "...EXPENS")
    #    is a liability, never a P&L expense, despite containing "exp".
    _BS_PARTY = ('creditor', 'debtor', 'payable', 'receivable',
                 'loan', 'borrow', 'provision', 'outstanding')
    is_bs_party = any(p in key for p in _BS_PARTY)
    # 3) Expense indicator: any "...Expenses/Charges/Fees/Bill" ledger is a P&L
    #    expense, never an asset or creditor. Guard runs before keyword guesses.
    #    ('prepaid expenses' is an asset group and only ever arrives via the
    #    Group column, which is matched before account names.)
    if not is_bs_party and 'prepaid' not in key \
            and any(ind in key for ind in _EXPENSE_INDICATORS):
        return 'INDIRECT EXPENSES'
    # 3) Standalone asset ledger names.
    if key in _ASSET_TERMS:
        return 'FIXED ASSETS'
    # 4) Keyword substring fallback.
    for keywords, target in _KEYWORD_MAP:
        if any(kw in key for kw in keywords):
            return target
    # 5) Retry once without a trailing "A/c" / "Account" suffix so ledgers like
    #    "Brokerage A/c" or "Depreciation A/c." still classify.
    import re
    stripped = re.sub(r'\s*(a/c\.?|account)\s*$', '', key).strip()
    if stripped and stripped != key:
        return _match(stripped)
    return None


def _normalise(group: str, account_name: str = '') -> str:
    # 1) Trust an explicit, recognised Group column first.
    matched = _match(group)
    if matched:
        return matched

    # 1b) If the group name itself is a section-header canonical name
    #     (set by Format J parser), resolve it directly.
    grp_upper = (group or '').strip().upper()
    if grp_upper in _SCHEDULE_SECTION_MAP:
        return _SCHEDULE_SECTION_MAP[grp_upper]

    # 2) No usable group → classify by the ledger/account name.
    matched = _match(account_name)
    if matched:
        return matched

    # 2b) Additional account-name heuristics for common unresolved cases:
    akey = (account_name or '').strip().lower()

    # Bank account names (HDFC Bank Ltd, ICICI Bank, SBI etc.)
    if any(b in akey for b in ('bank ltd', 'bank limited', ' bank', 'hdfc', 'icici',
                                'sbi', 'axis bank', 'kotak', 'yes bank', 'ubi',
                                'union bank', 'state bank', 'canara', 'pnb',
                                'punjab national', 'idbi', 'indusind')):
        # Bank OD/OCC → Secured Loans; savings/current → Cash and Bank
        if any(x in akey for x in ('od', 'occ', 'overdraft', 'cc a/c', 'cash credit')):
            return 'SECURED LOANS'
        return 'CASH AND BANK'

    # Partners Capital entries
    if 'partners capital' in akey or ('capital' in akey and 'partner' in akey):
        return 'CAPITAL'

    # "As Per Schedule X" — group comes from section context (already in grp_upper)
    # If group is already set to a valid canonical group, use it
    if akey.startswith('as per schedule') and grp_upper and grp_upper != 'UNGROUPED':
        return grp_upper

    # Person/party names under Sundry Creditors/Debtors context
    # (Group J parser sets group correctly; if it's already SUNDRY_*, use it)
    if grp_upper in ('SUNDRY CREDITORS', 'SUNDRY DEBTORS',
                     'LOANS AND ADVANCES (ASSETS)', 'UNSECURED LOANS',
                     'CAPITAL', 'PROVISIONS', 'CASH AND BANK',
                     'OTHER CURRENT ASSETS', 'FIXED ASSETS', 'INVESTMENTS',
                     'DUTIES AND TAXES', 'OPENING STOCK', 'CLOSING STOCK',
                     'DIRECT EXPENSES (M)', 'INDIRECT EXPENSES',
                     'INDIRECT INCOMES', 'SALES A/C', 'PURCHASE A/C',
                     'BALANCE WITH REVENUE AUTHORITY', 'DEPOSITS',
                     'MANUFACTURING EXPENSES', 'DIRECT INCOMES'):
        return grp_upper

    # 3) Keep an explicit-but-unknown group as-is; otherwise UNGROUPED.
    if group and group.strip():
        return group.strip().upper()
    return 'UNGROUPED'



class TrialBalanceCalculator:

    def compute(
        self, entries: List[TrialBalanceEntry]
    ) -> Dict[str, GroupSummary]:
        data = [e for e in entries if not e.is_total and not e.is_subtotal]
        groups: Dict[str, GroupSummary] = {}

        for entry in data:
            raw = entry.group or ''
            gname = _normalise(raw, entry.account_name)

            if gname not in groups:
                groups[gname] = GroupSummary(group_name=gname)
            grp = groups[gname]
            grp.entries.append(entry)
            grp.total_debit  += entry.debit
            grp.total_credit += entry.credit

        for grp in groups.values():
            grp.net_balance = grp.total_debit - grp.total_credit

        logger.info("Trial Balance: %d groups, %d entries", len(groups), len(data))
        return groups


# ── Profit & Loss group classification (mirrors the Summary sheet) ─────────────
_INCOME_GROUPS = {
    'SALES A/C', 'DIRECT INCOMES', 'INDIRECT INCOMES', 'INTEREST RECEIVED',
    'COMMISSION RECEIVED', 'RENT INCOME', 'DIVIDEND INCOME', 'OTHER INCOMES',
}
_EXPENSE_GROUPS = {
    'PURCHASE A/C', 'MANUFACTURING EXPENSES', 'DIRECT EXPENSES (M)', 'DIRECT WAGES',
    'POWER AND FUEL (M)', 'CARRIAGE INWARD', 'DIRECT EXPENSES', 'INDIRECT EXPENSES',
    'COMMISSION PAID', 'COMPENSATION TO EMPLOYEES', 'ADVERTISEMENT', 'FREIGHT',
    'INSURANCE', 'TELEPHONE', 'TRAVELLING', 'CONVEYANCE', 'RENT', 'RATES AND TAXES',
    'REPAIR & MAINTENANCE', 'POWER AND FUEL', 'DEPRECIATION', 'BAD DEBTS', 'DONATION',
    'OTHER EXPENSES', 'AUDITORS REMUNERATION', 'INTEREST PAID', 'FINANCIAL EXPENSES',
    'EXTRA-ORDINARY EXPENSES', 'STAFF WELFARE',
}


def compute_pl_net_profit(groups: Dict[str, GroupSummary]) -> float:
    """Net Profit / (Loss) derived ONLY from classified P&L groups.

    Positive = profit, negative = loss. This is the real result of operations —
    it is NOT the Dr/Cr difference of the trial balance. A trial balance that
    does not tie is a data error, never a profit.
    """
    income = expense = 0.0
    op_dr = op_cr = cl_dr = cl_cr = 0.0
    for name, g in groups.items():
        if name in _INCOME_GROUPS:
            income += g.total_credit - g.total_debit
        elif name in _EXPENSE_GROUPS:
            expense += g.total_debit - g.total_credit
        elif name == 'OPENING STOCK':
            op_dr += g.total_debit; op_cr += g.total_credit
        elif name == 'CLOSING STOCK':
            cl_dr += g.total_debit; cl_cr += g.total_credit
    # When BOTH the P&L (credit side) and the Balance Sheet (asset/debit
    # side) report the same stock figure, the two entries describe ONE
    # stock value — they must not cancel each other to zero.
    if cl_cr > 0 and cl_dr > 0 and abs(cl_cr - cl_dr) <= 1.0:
        closing = cl_cr
    else:
        closing = cl_cr - cl_dr
    if op_dr > 0 and op_cr > 0 and abs(op_dr - op_cr) <= 1.0:
        opening = op_dr
    else:
        opening = op_dr - op_cr
    return round(income - expense + closing - opening, 2)
