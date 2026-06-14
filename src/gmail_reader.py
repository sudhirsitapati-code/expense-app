"""
gmail_reader.py
Polls Gmail for SBI/ICICI bank alert emails and extracts expense details.
"""

import base64
import json
import os
import re
from datetime import datetime
from typing import Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
DATA_DIR = os.path.join(BASE_DIR, "data")

TOKEN_FILE = os.getenv("GMAIL_TOKEN_FILE", "config/gmail_token.json")
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

PROCESSED_IDS_PATH = os.path.join(DATA_DIR, "processed_gmail_ids.json")


def _get_service():
    creds = Credentials.from_authorized_user_file(
        os.path.join(BASE_DIR, TOKEN_FILE), SCOPES
    )
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return build("gmail", "v1", credentials=creds)


def _get_processed_ids() -> set:
    if not os.path.exists(PROCESSED_IDS_PATH):
        return set()
    with open(PROCESSED_IDS_PATH) as f:
        return set(json.load(f))


def _save_processed_id(msg_id: str):
    os.makedirs(DATA_DIR, exist_ok=True)
    ids = _get_processed_ids()
    ids.add(msg_id)
    with open(PROCESSED_IDS_PATH, "w") as f:
        json.dump(list(ids), f)


def _decode_body(msg) -> str:
    payload = msg.get("payload", {})
    parts = payload.get("parts", [])
    body = ""
    if parts:
        for part in parts:
            if part.get("mimeType") == "text/plain":
                data = part.get("body", {}).get("data", "")
                if data:
                    body = base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
                    break
    else:
        data = payload.get("body", {}).get("data", "")
        if data:
            body = base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
    return body


def _parse_sbi_alert(body: str) -> Optional[dict]:
    """
    SBI alert format:
    Dear Customer, Rs.XXXX.XX has been debited from account XX1234 on DD-MM-YY.
    Info: VENDOR NAME. Avl Bal: Rs.XXXXX.XX
    """
    amount_match = re.search(r"Rs\.?([\d,]+(?:\.\d{2})?)\s+has been debited", body, re.IGNORECASE)
    vendor_match = re.search(r"Info:\s*(.+?)(?:\.|Avl|Available|\n)", body, re.IGNORECASE)
    date_match = re.search(r"on\s+(\d{2}[-/]\d{2}[-/]\d{2,4})", body, re.IGNORECASE)

    if not amount_match:
        return None

    amount_str = amount_match.group(1).replace(",", "")
    vendor = vendor_match.group(1).strip() if vendor_match else "Unknown"
    date_str = date_match.group(1) if date_match else datetime.now().strftime("%d-%m-%y")

    return {
        "bank": "SBI",
        "amount": float(amount_str),
        "vendor": vendor,
        "date": date_str,
        "raw": body[:300],
    }


def _parse_icici_alert(body: str) -> Optional[dict]:
    """
    ICICI alert format:
    Dear Customer, INR XXXX.XX has been debited from your ICICI Bank Account XX1234
    on DD-MM-YYYY for VENDOR NAME. Available balance: INR XXXXX.XX
    """
    amount_match = re.search(r"INR\s*([\d,]+(?:\.\d{2})?)\s+has been debited", body, re.IGNORECASE)
    vendor_match = re.search(r"for\s+(.+?)(?:\.|Available|Avl|\n)", body, re.IGNORECASE)
    date_match = re.search(r"on\s+(\d{2}[-/]\d{2}[-/]\d{2,4})", body, re.IGNORECASE)

    if not amount_match:
        return None

    amount_str = amount_match.group(1).replace(",", "")
    vendor = vendor_match.group(1).strip() if vendor_match else "Unknown"
    date_str = date_match.group(1) if date_match else datetime.now().strftime("%d-%m-%y")

    return {
        "bank": "ICICI",
        "amount": float(amount_str),
        "vendor": vendor,
        "date": date_str,
        "raw": body[:300],
    }


def fetch_new_expenses() -> list[dict]:
    """Fetch unprocessed bank alert emails and return parsed expense dicts."""
    service = _get_service()
    processed_ids = _get_processed_ids()
    expenses = []

    sbi_label = os.getenv("SBI_GMAIL_LABEL", "SBI-Expenses")
    icici_label = os.getenv("ICICI_GMAIL_LABEL", "ICICI-Expenses")

    for label_name, parser in [(sbi_label, _parse_sbi_alert), (icici_label, _parse_icici_alert)]:
        try:
            labels_result = service.users().labels().list(userId="me").execute()
            label_id = next(
                (l["id"] for l in labels_result.get("labels", []) if l["name"] == label_name),
                None
            )
            if not label_id:
                continue

            messages_result = service.users().messages().list(
                userId="me", labelIds=[label_id], maxResults=20
            ).execute()

            for msg_ref in messages_result.get("messages", []):
                msg_id = msg_ref["id"]
                if msg_id in processed_ids:
                    continue

                msg = service.users().messages().get(
                    userId="me", id=msg_id, format="full"
                ).execute()

                body = _decode_body(msg)
                parsed = parser(body)

                if parsed:
                    parsed["gmail_id"] = msg_id
                    expenses.append(parsed)
                    _save_processed_id(msg_id)

        except Exception as e:
            print(f"Error fetching {label_name} emails: {e}")

    return expenses


if __name__ == "__main__":
    results = fetch_new_expenses()
    print(f"Found {len(results)} new expense(s):")
    for e in results:
        print(f"  {e['bank']} | Rs {e['amount']:,.0f} | {e['vendor']} | {e['date']}")
