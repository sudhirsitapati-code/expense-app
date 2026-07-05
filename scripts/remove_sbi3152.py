"""
Remove all master_ledger entries with account = SBI-3152 from production DB.
SBI-3152 is the same account as SBI-4852 — parser was picking up wrong digits.

Usage:
  cd /Users/sudhirsitapati/Desktop/expense-app
  python3 scripts/remove_sbi3152.py
"""

import os
import sys
import json
from dotenv import load_dotenv

load_dotenv()
import psycopg2

DATABASE_URL = os.getenv("DATABASE_URL", "").replace("postgresql://", "postgres://")
if not DATABASE_URL:
    print("ERROR: DATABASE_URL not set"); sys.exit(1)

conn = psycopg2.connect(DATABASE_URL)
cur = conn.cursor()

cur.execute("SELECT value FROM kv_store WHERE key = 'master_ledger'")
row = cur.fetchone()
if not row:
    print("No master_ledger found"); sys.exit(1)

ledger = row[0]
before = len(ledger)
bad = [t for t in ledger if (t.get("account") or "").upper() == "SBI-3152"]
ledger = [t for t in ledger if (t.get("account") or "").upper() != "SBI-3152"]
after = len(ledger)

print(f"Found {len(bad)} SBI-3152 entries to remove:")
for t in bad[:10]:
    print(f"  {t.get('date')} | {t.get('description','')[:50]} | {t.get('debit')}")
if len(bad) > 10:
    print(f"  ... and {len(bad)-10} more")

if not bad:
    print("Nothing to remove.")
    sys.exit(0)

confirm = input(f"\nRemove {len(bad)} entries ({before} → {after})? [y/N] ")
if confirm.strip().lower() != "y":
    print("Aborted."); sys.exit(0)

cur.execute(
    "UPDATE kv_store SET value = %s::jsonb, updated_at = NOW() WHERE key = 'master_ledger'",
    (json.dumps(ledger),)
)
conn.commit()
print(f"Done. Removed {before - after} entries.")
cur.close(); conn.close()
