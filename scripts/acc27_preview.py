"""
Preview ACC27 SudhirExpenses matching against the live Railway ledger.

Usage:
    python scripts/acc27_preview.py <railway-url> <session-cookie>

Example:
    python scripts/acc27_preview.py https://expense-app-xxx.up.railway.app "session=abc123"
"""
import sys, json, openpyxl, requests
from datetime import datetime

XL_PATH = "/Users/sudhirsitapati/Desktop/_ACC 27.xlsx"

HEADING_FIX = {
    'Staff salary': 'Staff Salary', 'Grocries': 'Groceries',
    'Internet': 'Misc', 'Miscellaneous': 'Misc',
}

def parse_date(v):
    if isinstance(v, datetime): return v
    if isinstance(v, str):
        for fmt in ('%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d'):
            try: return datetime.strptime(v.strip(), fmt)
            except: pass
    return None

def extract_xl():
    wb  = openpyxl.load_workbook(XL_PATH, data_only=True)
    ws  = wb['SudhirExpenses']
    out = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if row[0] is None: continue
        acct = str(row[2] or '').lower().replace(' ', '')
        if 'sbi' not in acct: continue
        d = parse_date(row[3])
        if not d: continue
        debit  = float(row[6]) if row[6] else 0
        credit = float(row[7]) if row[7] else 0
        amt = abs(debit or credit)
        if amt == 0: continue
        heading = HEADING_FIX.get(str(row[9] or '').strip(), str(row[9] or '').strip())
        out.append({
            'date': d.strftime('%d-%b-%y'),
            'account': 'SBI-4852',
            'debit': debit, 'credit': credit, 'amount': amt,
            'description': str(row[4] or '').strip(),
            'paid_to': str(row[5] or '').strip() or None,
            'type': str(row[8] or '').strip().lower(),
            'heading': heading,
            'notes': str(row[10] or '').strip() or None,
        })
    return out

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python scripts/acc27_preview.py <railway-url> <session-cookie>")
        sys.exit(1)

    base_url = sys.argv[1].rstrip('/')
    cookie   = sys.argv[2]

    entries = extract_xl()
    print(f"Extracted {len(entries)} SBI entries from Excel\n")

    # Send first 20 entries to get at least 1 good match for preview
    payload = {"entries": entries[:20], "limit": 1, "apply": False}
    r = requests.post(
        f"{base_url}/api/admin/acc27-match-preview",
        json=payload,
        headers={"Cookie": cookie},
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()

    matches = [x for x in data["results"] if x["match"] != "none"]
    if not matches:
        print("No matches found in first 20 entries. Try increasing the sample.")
        sys.exit(1)

    m = matches[0]
    print("=" * 60)
    print("EXAMPLE MATCH")
    print("=" * 60)
    print("\nEXCEL ENTRY:")
    xl = m["xl"]
    print(f"  Date:        {xl['date']}")
    print(f"  Account:     {xl['account']}")
    print(f"  Amount:      ₹{xl['amount']:,.0f} {'debit' if xl['debit'] else 'credit'}")
    print(f"  Description: {xl['description']}")
    print(f"  Type:        {xl['type']}")
    print(f"  Heading:     {xl['heading']}")
    if xl.get('notes'): print(f"  Notes:       {xl['notes']}")

    print(f"\nMATCHED LEDGER ENTRY (confidence: {m['match']}, {m['days_gap']}d gap):")
    ld = m["ledger"]
    print(f"  Seq:         #{ld.get('seq')}")
    print(f"  Date:        {ld.get('date')}")
    print(f"  Account:     {ld.get('account')}")
    print(f"  Debit:       ₹{ld.get('debit') or 0:,.0f}  Credit: ₹{ld.get('credit') or 0:,.0f}")
    print(f"  Raw desc:    {ld.get('raw_description')}")
    print(f"  Type now:    {ld.get('type')}  →  would become: {xl['type']}")
    print(f"  Heading now: {ld.get('heading')}  →  would become: {xl['heading']}")

    print("\nIf this looks correct, run with --apply to update all matches.")
    print(f"Total Excel entries to process: {data['total_xl']}")
