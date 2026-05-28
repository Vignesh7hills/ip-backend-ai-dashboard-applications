"""
API integration test — sends real PDF to the running server and saves output.
Run AFTER starting: uvicorn main:app --port 8000
"""
import sys
try:
    import requests
except ImportError:
    print("Install requests: pip install requests")
    sys.exit(1)

PDF = r'c:\Users\naniv\OneDrive\Desktop\sample files\loan_reporting_extracted\1. Loan Reporting\Sample Data\Ledger\MTA-LOAN AND ADVANCES 2024-25.pdf'
OUT = r'c:\Users\naniv\OneDrive\Desktop\api_test_output.xlsx'
BASE = 'http://localhost:8000'

print("Testing /health ...")
r = requests.get(f'{BASE}/health')
print(f"  {r.status_code}: {r.json()}")

print("\nTesting /api/loan-reporting-process ...")
with open(PDF, 'rb') as f:
    resp = requests.post(
        f'{BASE}/api/loan-reporting-process',
        files={'file': ('MTA-LOAN AND ADVANCES 2024-25.pdf', f, 'application/pdf')},
        data={'company_name': 'MAHABIR TEXTILE AGENCY'},
    )

print(f"  Status: {resp.status_code}")
print(f"  Records: {resp.headers.get('X-Records-Processed')}")
print(f"  Duration: {resp.headers.get('X-Duration-Ms')} ms")
print(f"  Warnings: {resp.headers.get('X-Warnings-Count')}")
print(f"  Content-Type: {resp.headers.get('Content-Type')}")

if resp.status_code == 200:
    with open(OUT, 'wb') as f:
        f.write(resp.content)
    print(f"\n  Excel saved: {OUT} ({len(resp.content):,} bytes)")
else:
    print(f"\n  ERROR: {resp.text}")
