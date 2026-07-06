"""
Trigger _sync_approvals_to_ledger() on production to populate missing entries.
Run AFTER the latest deploy is live.
"""
import os, sys
from pathlib import Path
from dotenv import load_dotenv
load_dotenv(Path(__file__).parent.parent / ".env")
import psycopg2, json
from collections import Counter
from datetime import datetime

DATABASE_URL = os.getenv("DATABASE_URL","").replace("postgresql://","postgres://")
if not DATABASE_URL:
    print("ERROR: DATABASE_URL not set"); sys.exit(1)

conn = psycopg2.connect(DATABASE_URL)
cur = conn.cursor()
cur.execute("SELECT value FROM kv_store WHERE key='master_ledger'")
ledger = cur.fetchone()[0]
cur.execute("SELECT value FROM kv_store WHERE key='approval_log'")
approval_log = cur.fetchone()[0]

existing_ids = {t["txn_id"] for t in ledger}

SBI_PMS = ("sbi","sbi-4852","sbi4852","sbi3152","sbi-3152","sbi-3142","sbi3142")

# Build a multiset of (YYYY-MM, amount) from existing SBI-4852 entries tagged
# as approval_log source — these are already confirmed by statement, so we skip
# the corresponding provisional entry to avoid showing duplicates.
from collections import defaultdict

def to_ym(date_str):
    """Return YYYY-MM from any date string format."""
    s = str(date_str or "")[:10]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d-%b-%Y"):
        try:
            from datetime import datetime as _dt
            return _dt.strptime(s, fmt).strftime("%Y-%m")
        except:
            pass
    return s[:7]

sbi_confirmed = defaultdict(int)
for t in ledger:
    acct = (t.get("account") or "").upper()
    if "4852" in acct and t.get("source") == "approval_log":
        ym = to_ym(t.get("date",""))
        amt = round(float(t.get("debit") or 0))
        sbi_confirmed[(ym, amt)] += 1

def parse_date(s):
    if not s: return ""
    return str(s)[:10]

def fy_info(date_str):
    try:
        d = datetime.strptime(date_str[:10], "%Y-%m-%d")
        if d.month >= 4:
            fy = d.year; mn = d.month - 3
        else:
            fy = d.year - 1; mn = d.month + 9
        months = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]
        return {"fy_year": f"FY{str(fy)[2:]}", "fy_month_no": mn, "fy_month_name": months[mn-1]}
    except:
        return {"fy_year":"", "fy_month_no":0, "fy_month_name":""}

APP_TO_HEADING = {
    "groceries":"Groceries","dining":"Dining","transport":"Transport",
    "utilities":"Utilities","entertainment":"Entertainment","shopping":"Shopping",
    "health":"Health","education":"Education","travel":"Travel",
    "maintenance":"Maintenance Expense","home_repair":"One Time Charge",
}
CANONICAL_HEADINGS = set(APP_TO_HEADING.values()) | {"Misc"}

to_add = []
skipped = []

for e in approval_log:
    if e.get("action") not in ("AUTO_APPROVE","APPROVED","APPROVED_LOWER"):
        continue
    pm = (e.get("payment_method") or "cash").lower()
    is_sbi = pm in SBI_PMS
    # upi/bank also sync immediately; only plain 'cash' needs confirmed_paid
    if not is_sbi and pm == "cash" and not e.get("confirmed_paid"):
        skipped.append(("cash_unconfirmed", e.get("request_id","?")))
        continue

    # Build txn_id (matches _approval_to_ledger_entry logic)
    txn_id = f"appr_{e.get('request_id','')}"
    if txn_id in existing_ids:
        continue

    # Skip SBI entries already confirmed by a statement entry (match by month+amount)
    if is_sbi:
        date_str_check = parse_date(e.get("timestamp",""))
        amt_check = round(float(e.get("approved_amount") or e.get("amount") or 0))
        key = (to_ym(date_str_check), amt_check)
        if sbi_confirmed.get(key, 0) > 0:
            sbi_confirmed[key] -= 1  # consume so we don't double-skip same amount
            skipped.append(("stmt_match", e.get("request_id","?")))
            continue

    acct  = "SBI-4852prov" if is_sbi else "cash"
    bank  = "SBI" if is_sbi else "approval"
    date_str = parse_date(e.get("timestamp",""))
    fy = fy_info(date_str)
    cat = e.get("category","")
    explicit_heading = e.get("heading","")
    heading = (explicit_heading if explicit_heading in CANONICAL_HEADINGS
               else cat if cat in CANONICAL_HEADINGS
               else APP_TO_HEADING.get(cat,"Misc"))

    txn = {
        "txn_id": txn_id,
        "date": date_str,
        "fy_month_no": fy["fy_month_no"],
        "fy_month_name": fy["fy_month_name"],
        "fy_year": fy["fy_year"],
        "account": acct,
        "account_type": pm,
        "bank": bank,
        "debit": float(e.get("approved_amount") or e.get("amount") or 0),
        "credit": 0,
        "paid_to": e.get("vendor",""),
        "heading": heading,
        "type": "personal",
        "source": "approval_log",
        "submitter": e.get("submitter",""),
        "request_id": e.get("request_id",""),
        "seq": 0,
    }
    to_add.append(txn)

if not to_add:
    print("Nothing to add — all approved entries already in ledger.")
    cur.close(); conn.close()
    sys.exit(0)

pm_breakdown = Counter(t["account_type"] for t in to_add)
print(f"\nWill add {len(to_add)} entries to master_ledger:")
for pm, cnt in pm_breakdown.most_common():
    acct = "SBI-4852prov" if pm in SBI_PMS else "cash"
    print(f"  {pm:<15} → {acct}  ({cnt})")
stmt_skipped = sum(1 for s in skipped if s[0] == "stmt_match")
cash_skipped = sum(1 for s in skipped if s[0] != "stmt_match")
if stmt_skipped:
    print(f"Skipping {stmt_skipped} SBI entries already confirmed by statement.")
if cash_skipped:
    print(f"Skipping {cash_skipped} unconfirmed cash entries.")

confirm = input("\nProceed? [y/N] ").strip().lower()
if confirm != "y":
    print("Aborted."); sys.exit(0)

ledger.extend(to_add)
ledger.sort(key=lambda t: t.get("date",""), reverse=True)

# Assign sequential seq numbers
seq = max((t.get("seq",0) for t in ledger if t.get("seq")), default=0)
for t in ledger:
    if not t.get("seq"):
        seq += 1
        t["seq"] = seq

cur.execute("UPDATE kv_store SET value=%s WHERE key='master_ledger'", [json.dumps(ledger)])
conn.commit()
cur.close(); conn.close()
print(f"\nDone — added {len(to_add)} entries.")
