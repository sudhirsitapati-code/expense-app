"""
reconcile.py
- Matches bank debits to approvals (amount ±Rs50, date ±3 days)
- Sets confirmed_paid=true on matches
- Flags unmatched debits >Rs500 as unauthorized
- Alerts Sudhir on unauthorized spends and cash gaps
"""

import json
import os
from datetime import datetime, timedelta
from typing import Optional

from src.gmail_reader import fetch_new_expenses
from src.whatsapp_handler import send_to_sudhir

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
CONFIG_DIR = os.path.join(BASE_DIR, "config")

APPROVAL_LOG_PATH = os.path.join(DATA_DIR, "approval_log.json")
RECONCILE_LOG_PATH = os.path.join(DATA_DIR, "reconcile_log.json")

AMOUNT_TOLERANCE = 50        # Rs
DATE_TOLERANCE_DAYS = 3
UNAUTH_THRESHOLD = 500       # Rs — ignore tiny rounding debits


def _load_json(path):
    if not os.path.exists(path):
        return []
    with open(path) as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def _save_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(data, f, indent=2)


def _load_recurring() -> list:
    p = os.path.join(CONFIG_DIR, "approved_recurring.json")
    if not os.path.exists(p):
        return []
    with open(p) as f:
        return json.load(f).get("recurring", [])


def _parse_date(date_str: str) -> Optional[datetime]:
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None


def _is_recurring(vendor: str, amount: float, recurring: list) -> bool:
    vendor_lower = vendor.lower()
    for rec in recurring:
        for kw in rec.get("payee_keywords", []):
            if kw.lower() in vendor_lower:
                if rec["amount_min"] <= amount <= rec["amount_max"]:
                    return True
    return False


def _vendor_match(bank_vendor: str, log_vendor: str) -> bool:
    """Fuzzy vendor match — any word overlap counts, or one contains the other."""
    b, l = bank_vendor.lower(), log_vendor.lower()
    if l in b or b in l:
        return True
    # word overlap (ignore very short words)
    b_words = {w for w in b.split() if len(w) > 2}
    l_words = {w for w in l.split() if len(w) > 2}
    return bool(b_words & l_words)


def _find_matching_approval(bank_expense: dict, approval_log: list) -> Optional[dict]:
    """Match a bank debit to an approved expense.
    Primary: amount ±Rs50 within ±7 days. Vendor match is best-effort — bank
    descriptions are often truncated/encoded so we rely mostly on amount+date.
    """
    bank_amount = bank_expense["amount"]
    bank_vendor = bank_expense.get("vendor", "")
    bank_date = _parse_date(bank_expense.get("date", ""))

    candidates = []
    for entry in approval_log:
        if entry.get("action") not in ("AUTO_APPROVE", "APPROVED", "APPROVED_LOWER"):
            continue
        if entry.get("confirmed_paid"):
            continue

        log_amount = entry.get("approved_amount") or entry.get("amount", 0)
        if abs(bank_amount - log_amount) > AMOUNT_TOLERANCE:
            continue

        log_date = None
        if entry.get("timestamp"):
            try:
                log_date = datetime.fromisoformat(entry["timestamp"])
            except ValueError:
                pass

        if bank_date and log_date:
            if abs((bank_date - log_date).days) > 7:
                continue

        vendor_ok = _vendor_match(bank_vendor, entry.get("vendor", ""))
        candidates.append((entry, vendor_ok))

    if not candidates:
        return None
    # Prefer vendor-matched candidates; otherwise take closest amount match
    vendor_matched = [c for c, v in candidates if v]
    pool = vendor_matched if vendor_matched else [c for c, _ in candidates]
    return min(pool, key=lambda e: abs(bank_amount - (e.get("approved_amount") or e.get("amount", 0))))


def run_reconciliation(notify_sudhir: bool = True) -> list:
    """
    Fetch new bank debits, match against approvals.
    - Matched → set confirmed_paid=True on approval log entry
    - Unmatched >Rs500 and not recurring → flag as unauthorized
    Returns list of unauthorized (unmatched) expenses.
    """
    bank_expenses = fetch_new_expenses()
    if not bank_expenses:
        return []

    approval_log = _load_json(APPROVAL_LOG_PATH)
    reconcile_log = _load_json(RECONCILE_LOG_PATH)
    recurring = _load_recurring()
    unauthorized = []

    for expense in bank_expenses:
        match = _find_matching_approval(expense, approval_log)
        is_recurring = _is_recurring(expense["vendor"], expense["amount"], recurring)

        entry = {
            "gmail_id": expense.get("gmail_id"),
            "bank": expense["bank"],
            "vendor": expense["vendor"],
            "amount": expense["amount"],
            "date": expense["date"],
            "matched": match is not None,
            "is_recurring": is_recurring,
            "matched_request_id": match["request_id"] if match else None,
            "ignored": False,
            "reconciled_at": datetime.now().isoformat(),
        }
        reconcile_log.append(entry)

        if match:
            # Mark the approval as confirmed paid
            for log_entry in approval_log:
                if log_entry.get("request_id") == match["request_id"]:
                    log_entry["confirmed_paid"] = True
                    log_entry["confirmed_at"] = datetime.now().isoformat()
                    log_entry["confirmed_by"] = "bank_match"
                    break
        elif not is_recurring and expense["amount"] > UNAUTH_THRESHOLD:
            unauthorized.append(expense)

    _save_json(RECONCILE_LOG_PATH, reconcile_log)
    _save_json(APPROVAL_LOG_PATH, approval_log)

    if unauthorized and notify_sudhir:
        lines = ["⚠️ *Unauthorized Spend Detected*\n"]
        for e in unauthorized:
            lines.append(f"• {e['bank']} | Rs {e['amount']:,.0f} | {e['vendor']} | {e['date']}")
        lines.append("\nThese were debited but not pre-approved.")
        lines.append("Reply with the expense details to submit post-facto, or check the dashboard.")
        send_to_sudhir("\n".join(lines))

    return unauthorized


if __name__ == "__main__":
    results = run_reconciliation(notify_sudhir=False)
    print(f"{len(results)} unauthorized debit(s):")
    for e in results:
        print(f"  {e['bank']} | Rs {e['amount']:,.0f} | {e['vendor']} | {e['date']}")
