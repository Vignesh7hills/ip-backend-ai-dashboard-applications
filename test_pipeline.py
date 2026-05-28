"""Quick end-to-end pipeline test — run from backend/ directory."""
import sys
sys.path.insert(0, r'c:\Users\naniv\OneDrive\Desktop\backend')

from app.modules.loan_reporting.service import LoanReportingService

PDF = r'c:\Users\naniv\OneDrive\Desktop\sample files\loan_reporting_extracted\1. Loan Reporting\Sample Data\Ledger\MTA-LOAN AND ADVANCES 2024-25.pdf'
OUT = r'c:\Users\naniv\OneDrive\Desktop\test_loan_reporting.xlsx'

svc = LoanReportingService()
result = svc.process(PDF, company_name='MAHABIR TEXTILE AGENCY')

print(f"Records  : {result['records']}")
print(f"Duration : {result['duration_ms']:.0f} ms")
print(f"Errors   : {len(result['errors'])}")
print(f"Warnings : {len(result['warnings'])}")
for e in result['errors'][:10]:
    print(f"  ERROR: {e}")
for w in result['warnings'][:5]:
    print(f"  WARN : {w}")

with open(OUT, 'wb') as f:
    f.write(result['excel_bytes'])
print(f"\nExcel saved: {OUT}")
