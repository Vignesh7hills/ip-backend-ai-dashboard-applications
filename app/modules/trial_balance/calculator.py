"""
Trial Balance Calculator with intelligent group normalisation.
Maps raw parsed group names → canonical Genius Software group names.
"""
from typing import List, Dict, Optional
from dataclasses import dataclass, field
from app.modules.trial_balance.parser import TrialBalanceEntry
from app.core.logger import get_logger

logger = get_logger(__name__)


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
    # FIX: CURRENT LIABILITIES group in BS format maps to PROVISIONS (liability side)
    'current liabilities':             'PROVISIONS',

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

    # Sales
    'sales a/c':                       'SALES A/C',
    'sale a/c':                        'SALES A/C',
    'sales':                           'SALES A/C',
    'sale account':                    'SALES A/C',
    'sale account':                    'SALES A/C',

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
    (['sundry creditor', 'karkhandar', 'broker'],
     'SUNDRY CREDITORS'),
    (['sundry debtor'],
     'SUNDRY DEBTORS'),
    (['fixed asset', 'plant', 'machinery', 'furniture', 'equipment',
      'building', 'vehicle', 'computer'],
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
    (['loans and advance', 'loans & advance'],
     'LOANS AND ADVANCES (ASSETS)'),
    (['manufacture', 'manufacturing'],
     'MANUFACTURING EXPENSES'),
    (['purchase'],
     'PURCHASE A/C'),
    (['sale'],
     'SALES A/C'),
    (['ganesh', 'ganeshji'],
     'CAPITAL'),
]


def _match(text: str) -> Optional[str]:
    """Return the canonical group for a single string, or None if unknown."""
    if not text:
        return None
    key = text.strip().lower()
    if key in _ALIASES:
        return _ALIASES[key]
    for keywords, target in _KEYWORD_MAP:
        if any(kw in key for kw in keywords):
            return target
    return None


def _normalise(group: str, account_name: str = '') -> str:
    # 1) Trust an explicit, recognised Group column first.
    matched = _match(group)
    if matched:
        return matched
    # 2) No usable group (e.g. a PDF/CSV with no Group column) →
    #    classify by the ledger/account name instead of dumping to UNGROUPED.
    matched = _match(account_name)
    if matched:
        return matched
    # 3) Keep an explicit-but-unknown group as-is; otherwise leave ungrouped.
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
    closing = cl_cr - cl_dr
    opening = op_dr - op_cr
    return round(income - expense + closing - opening, 2)
