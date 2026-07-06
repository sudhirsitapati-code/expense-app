"""
Check June Vincent SBI entries vs SBI-4852 statement entries.
Strict matching: amount within 2% AND ±₹50, date within 5 days.
"""
import os, sys, json
from datetime import datetime, date
from dotenv import load_dotenv
load_dotenv()
import psycopg2

DATABASE_URL = os.getenv("DATABASE_URL","").replace("postgresql://","postgres://")
if not DATABASE_URL:
    print("ERROR: DATABASE_URL not set"); sys.exit(1)

conn = psycopg2.connect(DATABASE_URL)
cur = conn.cursor()
cur.execute("SELECT value FROM kv_store WHERE key='master_ledger'")
ledger = cur.fetchone()[0]
cur.execute("SELECT value FROM kv_store WHERE key='approval_log'")
approval_log = cur.fetchone()[0]
cur.close(); conn.close()

JUN_START = date(2026, 6, 1)
JUN_END   = date(2026, 6, 30)

def parse_date(s):
    if not s: return None
    for fmt in ("%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%d-%b-%Y","%d %b %Y"):
        try: return datetime.strptime(str(s)[:10], fmt).date()
        except: pass
    return None

def amounts_match(a, b):
    """True if amounts are within 5% AND within ₹500."""
    diff = abs(a - b)
    pct  = diff / max(a, b, 1)
    return diff <= 500 and pct <= 0.05

SBI_PMS = ("sbi","sbi-4852","sbi4852","sbi3152","sbi-3152","sbi-3142","sbi3142")
vincent_sbi = []
for e in approval_log:
    if e.get("action") not in ("AUTO_APPROVE","APPROVED","APPROVED_LOWER"):
        continue
    if (e.get("submitter") or "").lower() != "vincent":
        continue
    pm = (e.get("payment_method") or "").lower()
    if pm not in SBI_PMS:
        continue
    dt = parse_date((e.get("timestamp") or "")[:10])
    if not dt or not (JUN_START <= dt <= JUN_END):
        continue
    vincent_sbi.append({
        "id":     e.get("request_id","?"),
        "date":   dt,
        "vendor": e.get("vendor","?"),
        "amount": float(e.get("approved_amount") or e.get("amount") or 0),
    })

sbi_statement = []
for t in ledger:
    acct = (t.get("account") or "").upper()
    if "4852" not in acct and "3152" not in acct:
        continue
    if t.get("source") == "approval_log":
        continue
    debit = float(t.get("debit") or 0)
    if debit == 0:
        continue
    dt = parse_date(t.get("date",""))
    if not dt or not (JUN_START <= dt <= JUN_END):
        continue
    sbi_statement.append({
        "txn_id": t.get("txn_id","?"),
        "date":   dt,
        "desc":   (t.get("paid_to") or t.get("transaction_details") or t.get("raw_description") or "")[:50],
        "amount": debit,
    })

print(f"\n{'='*65}")
print(f"Vincent SBI approvals in June:      {len(vincent_sbi)}")
print(f"SBI-4852 statement entries in June: {len(sbi_statement)}")
print(f"{'='*65}\n")

matched_stmt_ids = set()
matched_vincent  = []
unmatched_vincent = []

for v in sorted(vincent_sbi, key=lambda x: x["date"]):
    best = None
    best_days = 999
    for s in sbi_statement:
        if s["txn_id"] in matched_stmt_ids:
            continue
        if not amounts_match(v["amount"], s["amount"]):
            continue
        days = abs((s["date"] - v["date"]).days)
        if days > 5:
            continue
        if days < best_days:
            best = s
            best_days = days
    if best:
        matched_stmt_ids.add(best["txn_id"])
        matched_vincent.append(v)
        print(f"  ✓  {v['date']} ₹{v['amount']:>10,.0f}  {v['vendor'][:30]:<30}")
        print(f"      → {best['date']} ₹{best['amount']:>10,.0f}  {best['desc'][:40]}")
    else:
        unmatched_vincent.append(v)
        print(f"  ✗  {v['date']} ₹{v['amount']:>10,.0f}  {v['vendor'][:30]}  ← NO STMT MATCH")

print(f"\n{'='*65}")
print(f"Matched:   {len(matched_vincent)} / {len(vincent_sbi)}")
print(f"No match:  {len(unmatched_vincent)}  ← these need SBI-4852prov")
print(f"{'='*65}\n")
if unmatched_vincent:
    print("Vincent entries with no statement match:")
    for v in unmatched_vincent:
        print(f"  {v['date']}  ₹{v['amount']:,.0f}  {v['vendor']}")
