import os, sys
from pathlib import Path
from dotenv import load_dotenv
load_dotenv(Path(__file__).parent.parent / ".env")
import psycopg2
from collections import Counter

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

pm_counts = Counter((e.get("payment_method") or "?").lower() for e in approval_log
                    if e.get("action") in ("AUTO_APPROVE","APPROVED","APPROVED_LOWER"))
print("Approval log entries by payment_method:")
for pm, count in pm_counts.most_common():
    print(f"  {pm:<20} {count}")

print("\nMaster ledger approval_log source entries by account:")
appr_entries = [t for t in ledger if t.get("source") == "approval_log"]
acct_counts = Counter(t.get("account","?") for t in appr_entries)
for acct, count in sorted(acct_counts.items(), key=lambda x: -x[1]):
    print(f"  {acct:<25} {count}")

print("\nSample txn_ids of existing approval_log source entries:")
for t in appr_entries[:5]:
    print(f"  {t.get('txn_id','?')[:40]:<40}  acct={t.get('account')}  {t.get('paid_to','')[:25]}")

# How many approved entries have their request_id referenced in the ledger?
ledger_request_ids = {t.get("request_id") for t in ledger if t.get("request_id")}
ledger_appr_txnids = {t["txn_id"] for t in ledger if t.get("txn_id","").startswith("appr_")}
approved = [e for e in approval_log if e.get("action") in ("AUTO_APPROVE","APPROVED","APPROVED_LOWER")]
matched_by_reqid  = [e for e in approved if e.get("request_id") in ledger_request_ids]
matched_by_txnid  = [e for e in approved if f"appr_{e.get('request_id','')}" in ledger_appr_txnids]
truly_missing     = [e for e in approved
                     if e.get("request_id") not in ledger_request_ids
                     and f"appr_{e.get('request_id','')}" not in ledger_appr_txnids]

print(f"\nApproved entries matched by request_id in ledger: {len(matched_by_reqid)}")
print(f"Approved entries matched by appr_ txn_id:         {len(matched_by_txnid)}")
print(f"Truly missing (not in ledger at all):              {len(truly_missing)}")

pm_missing = Counter((e.get("payment_method") or "?").lower() for e in truly_missing)
print("\nMissing entries by payment_method:")
for pm, count in pm_missing.most_common():
    print(f"  {pm:<20} {count}")

conf  = sum(1 for e in approval_log if e.get("confirmed_paid") and e.get("action") in ("AUTO_APPROVE","APPROVED","APPROVED_LOWER"))
noconf = sum(1 for e in approval_log if not e.get("confirmed_paid") and e.get("action") in ("AUTO_APPROVE","APPROVED","APPROVED_LOWER"))
print(f"\nconfirmed_paid=True:  {conf}")
print(f"confirmed_paid=False: {noconf}")
print(f"\nTotal approval_log entries in master_ledger: {len(appr_entries)}")
print(f"Total approved approval_log entries:         {conf + noconf}")
