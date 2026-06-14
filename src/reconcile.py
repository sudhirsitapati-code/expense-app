"""
reconcile.py
Matches Gmail bank debit alerts against the approval log to flag unsubmitted expenses.
"""

import json
import os
from datetime import datetime
from typing import Optional

from src.gmail_reader import fetch_new_expenses
from src.whatsapp_handler import send_to_sudhir

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")

APPROVAL_LOG_PATH = os.path.join(DATA_DIR, "approval_log.json")
RECONCILE_LOG_PATH = os.path.join(DATA_DIR, "reconcile_log.json")

AMOUNT_TOLERANCE = 0.10  # 10% tolerance for matching


def _load_approval_log() -> list:
    if not os.path.exists(APPROVAL_LOG_PATH):
        return []
    with open(APPROVAL_LOG_PATH) as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def _load_reconcile_log() -> list:
    if not os.path.exists(RECONCILE_LOG_PATH):
        return []
    with open(RECONCILE_LOG_PATH) as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def _save_reconcile_log(log: list):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(RECONCILE_LOG_PATH, "w") as f:
        json.dump(log, f, indent=2)


def _find_matching_approval(bank_expense: dict, approval_log: list) -> Optional[dict]:
    """Try to match a bank debit to an approved/auto-approved expense."""
    bank_amount = bank_expense["amount"]
    bank_vendor = bank_expense["vendor"].lower()

    for entry in approval_log:
        if entry.get("action") not in ("AUTO_APPROVE", "APPROVED", "APPROVED_LOWER"):
            continue

        log_amount = entry.get("approved_amount") or entry.get("amount", 0)
        log_vendor = entry.get("vendor", "").lower()

        amount_diff = abs(bank_amount - log_amount) / max(log_amount, 1)
        vendor_words = set(bank_vendor.split()) & set(log_vendor.split())

        if amount_diff <= AMOUNT_TOLERANCE and (vendor_words or bank_vendor in log_vendor or log_vendor in bank_vendor):
            return entry

    return None


def run_reconciliation(notify_sudhir: bool = True) -> list:
    """
    Fetch new bank debits and check each against the approval log.
    Returns list of unmatched (unsubmitted) expenses.
    """
    bank_expenses = fetch_new_expenses()
    if not bank_expenses:
        return []

    approval_log = _load_approval_log()
    reconcile_log = _load_reconcile_log()
    unmatched = []

    for expense in bank_expenses:
        match = _find_matching_approval(expense, approval_log)
        entry = {
            "gmail_id": expense.get("gmail_id"),
            "bank": expense["bank"],
            "vendor": expense["vendor"],
            "amount": expense["amount"],
            "date": expense["date"],
            "matched": match is not None,
            "matched_request_id": match["request_id"] if match else None,
            "reconciled_at": datetime.now().isoformat(),
        }
        reconcile_log.append(entry)

        if not match:
            unmatched.append(expense)

    _save_reconcile_log(reconcile_log)

    if unmatched and notify_sudhir:
        lines = [f"⚠️ *Unsubmitted Bank Debits Detected*\n"]
        for e in unmatched:
            lines.append(f"• {e['bank']} | Rs {e['amount']:,.0f} | {e['vendor']} | {e['date']}")
        lines.append("\nThese were debited but not submitted for approval.")
        send_to_sudhir("\n".join(lines))

    return unmatched


if __name__ == "__main__":
    unmatched = run_reconciliation(notify_sudhir=False)
    print(f"{len(unmatched)} unmatched debit(s):")
    for e in unmatched:
        print(f"  {e['bank']} | Rs {e['amount']:,.0f} | {e['vendor']} | {e['date']}")
