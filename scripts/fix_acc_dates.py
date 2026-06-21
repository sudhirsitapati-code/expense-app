"""
fix_acc_dates.py
Corrects dates for all acc24/acc25 imported ledger entries.
The original import used raw Excel dates which were unreliable.
This script derives the correct date as the 1st of the correct calendar month
from fy_month_no + fy_year stored on each entry.

Usage:
    DATABASE_URL="..." python3 scripts/fix_acc_dates.py          # dry-run
    DATABASE_URL="..." python3 scripts/fix_acc_dates.py --apply  # apply fixes
"""

import sys
import os

sys.path.insert(0, '/Users/sudhirsitapati/Desktop/expense-app')

import src.db as _db

LEDGER_PATH = 'master_ledger'

# FY month_no → calendar month number (1=Apr ... 12=Mar)
FY_MONTH_TO_CAL = {
    1: 4, 2: 5, 3: 6, 4: 7, 5: 8, 6: 9,
    7: 10, 8: 11, 9: 12, 10: 1, 11: 2, 12: 3,
}

CAL_ABBR = {
    1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
    7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec',
}


def correct_date(month_no, fy_year):
    """Return date string like '01-Apr-23' derived from month_no + fy_year."""
    cal_month = FY_MONTH_TO_CAL[month_no]
    cal_year = fy_year - 1 if month_no <= 9 else fy_year
    year_2d = str(cal_year)[-2:]
    return f"01-{CAL_ABBR[cal_month]}-{year_2d}"


def main():
    apply = '--apply' in sys.argv

    print("Loading master ledger from DB...")
    transactions = _db.load(LEDGER_PATH, default=[])
    if not transactions:
        print("No data found.")
        return
    if isinstance(transactions, dict):
        transactions = transactions.get('transactions', [])
    print(f"Total entries: {len(transactions)}")

    acc_entries = [t for t in transactions if t.get('source') in ('acc24', 'acc25')]
    print(f"acc24/acc25 entries: {len(acc_entries)}")

    fixed = 0
    errors = []
    for t in acc_entries:
        month_no = t.get('fy_month_no')
        fy_year = t.get('fy_year')
        if not month_no or not fy_year or month_no not in FY_MONTH_TO_CAL:
            errors.append(t.get('txn_id'))
            continue

        correct = correct_date(int(month_no), int(fy_year))
        old = t.get('date', '')
        if old != correct:
            if not apply:
                if fixed < 5:
                    print(f"  WOULD FIX {t['txn_id']}: {old!r} → {correct!r}  (source={t['source']} month={month_no} fy={fy_year})")
            t['date'] = correct
            fixed += 1

    print(f"\nEntries that need fixing: {fixed}")
    if errors:
        print(f"Entries with missing month_no/fy_year (skipped): {len(errors)}")

    if apply:
        print("Saving back to DB...")
        _db.save(LEDGER_PATH, transactions)
        print(f"Done. Fixed {fixed} dates.")
    else:
        print("\nDry run — pass --apply to write changes.")


if __name__ == '__main__':
    main()
