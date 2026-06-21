"""
import_acc_transactions.py
Parse acc24.xlsx and acc25.xlsx and insert transactions into the master ledger.
Usage:
    python scripts/import_acc_transactions.py           # dry-run (print first 10)
    python scripts/import_acc_transactions.py --insert  # actual insert
"""

import sys
import os
import hashlib
from datetime import datetime, date

sys.path.insert(0, '/Users/sudhirsitapati/Desktop/expense-app')

import openpyxl

# ── Config ────────────────────────────────────────────────────────────────────

FILES = [
    {
        'path': '/Users/sudhirsitapati/Desktop/expense-app/data/FY24/acc24.xlsx',
        'sheet': 'SudhirExpenses',
        'fy_year': 2024,
        'source': 'acc24',
        # col indices (0-based): month_no, month_name, account, date, description,
        #   paid_to, debit, credit, net, type, heading, remarks
        'cols': {
            'month_no': 0, 'month_name': 1, 'account': 2, 'date': 3,
            'description': 4, 'paid_to': 5, 'debit': 6, 'credit': 7,
            'net': 8, 'type': 9, 'heading': 10, 'remarks': 11,
        },
    },
    {
        'path': '/Users/sudhirsitapati/Desktop/expense-app/data/FY25/ACC 25.xlsx',
        'sheet': 'SudhirExpenses',
        'fy_year': 2025,
        'source': 'acc25',
        # col indices (0-based): month_no, month_name, account, date, description,
        #   paid_to, debit, credit, type, None, heading, None, remarks
        'cols': {
            'month_no': 0, 'month_name': 1, 'account': 2, 'date': 3,
            'description': 4, 'paid_to': 5, 'debit': 6, 'credit': 7,
            'type': 8, 'heading': 10, 'remarks': 12,
        },
    },
]

HEADING_MAP = {
    'staff salary': 'Staff Salary',
    'maintenance expense': 'Maintenance Expense',
    'maintenance expenses': 'Maintenance Expense',
    'maintanence expenses': 'Maintenance Expense',
    'maintanence expense': 'Maintenance Expense',
    'monthly expense': 'Groceries',
    'monthly expenses': 'Groceries',
    'groceries': 'Groceries',
    'clothing': 'Clothes',
    'clothes': 'Clothes',
    'financial expense': 'Financial Expense / OD Interest',
    'financial expenses': 'Financial Expense / OD Interest',
    'od interest': 'Financial Expense / OD Interest',
    'c71': 'Home office',
    'c 71': 'Home office',
    'home office': 'Home office',
    'c51': 'Amma',
    'c 51': 'Amma',
    'a131': 'Ketki',
    'a 131': 'Ketki',
    'children education': 'Children Education',
    'education': 'Children Education',
    'home loan': 'Home Loan',
    'insurance': 'Insurance',
    'holiday': 'Holiday',
    'eating out': 'Eating Out',
    'alcohol': 'Alcohol',
    'medical': 'Medical',
    'gifts': 'Gifts',
    'gift': 'Gifts',
    'misc': 'Misc',
    'cash': 'Cash',
    'electricity': 'Electricity & Gas',
    'gas': 'Electricity & Gas',
    'electricity & gas': 'Electricity & Gas',
    'kalpataru': 'Kalpataru Maintenance',
    'kalpataru maintenance': 'Kalpataru Maintenance',
    'rent': 'Rent',
    'personal loan': 'Personal Loans',
    'personal loans': 'Personal Loans',
    'club': 'Club',
    'kashid': 'Kashid',
    'tax': 'Tax',
    'advance tax': 'Tax',
    'charity': 'Charity',
    'malhar': 'Malhar',
    'wellness': 'Wellness',
    'entertainment': 'Entertainment',
    'one time charge': 'One Time Charge',
    'uspaar': 'Uspaar',
    'amma': 'Amma',
    'ketki': 'Ketki',
    'interbank': 'Interbank',
    'salary': 'Salary',
    'interest': 'Interest',
    'equity': 'Equity',
    'loan': 'Loan',
}

SKIP_HEADINGS = {'Interbank', 'Salary', 'Interest', 'Equity', 'Loan'}

# FY month_no → calendar month (1=Apr ... 12=Mar)
FY_MONTH_TO_CAL = {
    1: 4, 2: 5, 3: 6, 4: 7, 5: 8, 6: 9,
    7: 10, 8: 11, 9: 12, 10: 1, 11: 2, 12: 3,
}

FY_MONTH_NAMES = {
    1: 'Apr', 2: 'May', 3: 'Jun', 4: 'Jul', 5: 'Aug', 6: 'Sep',
    7: 'Oct', 8: 'Nov', 9: 'Dec', 10: 'Jan', 11: 'Feb', 12: 'Mar',
}


def normalize_account(raw):
    if not raw:
        return None, None
    s = str(raw).strip().upper().replace(' ', '').replace('-', '')
    if 'SBI' in s and '4852' in s:
        return 'SBI-4852', 'SBI'
    if 'SBI' in s:
        return 'SBI-4852', 'SBI'
    if '1331' in s:
        return 'ICICI-1331', 'ICICI'
    if '0018' in s:
        return 'ICICI-0018', 'ICICI'
    if '9175' in s or ('975' in s and 'ICIC' in s):
        return 'ICICI-9175', 'ICICI'
    if 'ICIC' in s:
        return s, 'ICICI'
    return raw.strip(), None


def normalize_heading(raw):
    if not raw:
        return ''
    key = str(raw).strip().lower()
    return HEADING_MAP.get(key, str(raw).strip())


def normalize_type(raw):
    if not raw:
        return 'expense'
    s = str(raw).strip().lower()
    if s == 'error':
        return 'ERROR'
    return s if s in ('expense', 'income', 'transfer', 'investment') else 'expense'


def to_float(v):
    if v is None:
        return 0.0
    try:
        f = float(v)
        return f if f > 0 else 0.0
    except (ValueError, TypeError):
        return 0.0


def parse_date(val, month_no, fy_year):
    """Return (date_obj, used_fallback)"""
    d = None
    if val is not None:
        if isinstance(val, (datetime, date)):
            d = val if isinstance(val, datetime) else datetime(val.year, val.month, val.day)
        elif isinstance(val, str):
            val = val.strip()
            for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y', '%d-%m-%y'):
                try:
                    d = datetime.strptime(val, fmt)
                    break
                except ValueError:
                    pass

    if d and (2020 <= d.year <= 2026):
        return d, False

    # Fallback: derive from month_no + fy_year
    cal_month = FY_MONTH_TO_CAL[month_no]
    # For FY24: months 1-9 → 2023, months 10-12 → 2024
    # For FY25: months 1-9 → 2024, months 10-12 → 2025
    if month_no <= 9:
        cal_year = fy_year - 1
    else:
        cal_year = fy_year
    return datetime(cal_year, cal_month, 1), True


def txn_id(date_str, account, description, amount):
    raw = f"{date_str}|{account}|{description}|{amount}"
    return hashlib.sha1(raw.encode()).hexdigest()[:16]


def get_cell(row, idx):
    """Safely get cell value from a row tuple (0-based index)."""
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def parse_file(cfg):
    wb = openpyxl.load_workbook(cfg['path'], data_only=True)
    ws = wb[cfg['sheet']]
    cols = cfg['cols']
    fy_year = cfg['fy_year']
    source = cfg['source']

    transactions = []
    skip_counts = {'empty': 0, 'error_type': 0, 'skip_heading': 0, 'bad_month': 0}
    row_num = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        row_num += 1

        month_no_raw = get_cell(row, cols['month_no'])
        description = get_cell(row, cols['description'])
        debit = to_float(get_cell(row, cols['debit']))
        credit = to_float(get_cell(row, cols['credit']))

        # Skip empty rows
        desc_str = str(description).strip() if description else ''
        if not desc_str and debit == 0 and credit == 0:
            skip_counts['empty'] += 1
            continue

        # Parse month_no
        try:
            month_no = int(float(str(month_no_raw))) if month_no_raw else None
        except (ValueError, TypeError):
            month_no = None

        if not month_no or month_no not in FY_MONTH_TO_CAL:
            skip_counts['bad_month'] += 1
            continue

        # Type
        txn_type_raw = get_cell(row, cols['type'])
        txn_type = normalize_type(txn_type_raw)
        if txn_type == 'ERROR':
            skip_counts['error_type'] += 1
            continue

        # Heading
        heading_raw = get_cell(row, cols['heading'])
        heading = normalize_heading(heading_raw)
        if heading in SKIP_HEADINGS:
            skip_counts['skip_heading'] += 1
            continue

        # Date
        date_val = get_cell(row, cols['date'])
        date_obj, used_fallback = parse_date(date_val, month_no, fy_year)
        date_str = date_obj.strftime('%d-%b-%y')

        # Account
        account_raw = get_cell(row, cols['account'])
        account, bank = normalize_account(account_raw)

        # month_name
        month_name = FY_MONTH_NAMES[month_no]

        # paid_to
        paid_to = str(get_cell(row, cols['paid_to']) or '').strip()

        # remarks
        remarks_idx = cols.get('remarks')
        remarks = str(get_cell(row, remarks_idx) or '').strip() if remarks_idx is not None else ''

        # Amount for txn_id
        amount = debit if debit > 0 else credit

        tid = txn_id(date_str, account or '', desc_str, amount)

        txn = {
            'txn_id': tid,
            'date': date_str,
            'fy_month_no': month_no,
            'fy_month_name': month_name,
            'fy_year': fy_year,
            'account': account or str(account_raw or '').strip(),
            'account_type': 'savings',
            'bank': bank or '',
            'raw_description': desc_str,
            'paid_to': paid_to,
            'debit': debit,
            'credit': credit,
            'type': txn_type,
            'heading': heading,
            'remarks': remarks,
            'uncertain': False,
            'uncertain_fields': [],
            'confidence': 'acc_import',
            'source': source,
            'gmail_id': None,
            'ai_saving_tip': None,
            'saving_agreed': None,
            'reconciled_with': None,
            'created_at': datetime.now().isoformat(),
        }
        transactions.append(txn)

    return transactions, skip_counts, row_num


def main():
    dry_run = '--insert' not in sys.argv

    all_transactions = []
    total_rows = 0
    total_skips = {'empty': 0, 'error_type': 0, 'skip_heading': 0, 'bad_month': 0}

    for cfg in FILES:
        print(f"\nParsing {cfg['source']}: {cfg['path']}")
        txns, skips, rows = parse_file(cfg)
        print(f"  Rows scanned: {rows}, Parsed: {len(txns)}, Skipped: {sum(skips.values())}")
        for k, v in skips.items():
            if v:
                print(f"    - {k}: {v}")
        all_transactions.extend(txns)
        total_rows += rows
        for k in total_skips:
            total_skips[k] += skips[k]

    print(f"\nTotal parsed transactions: {len(all_transactions)}")

    if dry_run:
        print("\n=== DRY RUN: First 10 transactions ===")
        for t in all_transactions[:10]:
            print(f"  [{t['source']}] {t['date']} | {t['account']} | {t['raw_description'][:40]} | "
                  f"Dr:{t['debit']} Cr:{t['credit']} | {t['type']} | {t['heading']} | id:{t['txn_id']}")
        print("\nRun with --insert to actually save to DB.")
        return

    # Load existing ledger
    print("\nLoading master ledger from DB...")
    from src.master_ledger import _load_json, _save_json, LEDGER_PATH
    ledger = _load_json(LEDGER_PATH)
    existing_ids = {t['txn_id'] for t in ledger if t.get('txn_id')}
    print(f"Existing transactions: {len(ledger)}")

    inserted = 0
    duplicates = 0
    for t in all_transactions:
        if t['txn_id'] in existing_ids:
            duplicates += 1
        else:
            ledger.append(t)
            existing_ids.add(t['txn_id'])
            inserted += 1

    print(f"Inserting {inserted} new transactions ({duplicates} duplicates skipped)...")
    _save_json(LEDGER_PATH, ledger)
    print("Done.")

    print(f"\n=== SUMMARY ===")
    print(f"Total rows scanned:    {total_rows}")
    print(f"Skipped (empty):       {total_skips['empty']}")
    print(f"Skipped (error type):  {total_skips['error_type']}")
    print(f"Skipped (heading):     {total_skips['skip_heading']}")
    print(f"Skipped (bad month):   {total_skips['bad_month']}")
    print(f"Parsed:                {len(all_transactions)}")
    print(f"Duplicates skipped:    {duplicates}")
    print(f"Inserted:              {inserted}")
    print(f"Ledger size now:       {len(ledger)}")


if __name__ == '__main__':
    main()
