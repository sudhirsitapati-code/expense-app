"""
icici_statement_parser.py
Fetches ICICI statement PDFs from Gmail, decrypts them, parses all transactions,
and saves to data/icici_transactions.json.
"""

import base64
import io
import json
import os
import re
from datetime import datetime
from typing import Optional

import pdfplumber
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
CONFIG_DIR = os.path.join(BASE_DIR, "config")

TOKEN_FILE = os.path.join(BASE_DIR, os.getenv("GMAIL_TOKEN_FILE", "config/gmail_token.json"))
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
ICICI_LABEL = os.getenv("ICICI_GMAIL_LABEL", "ICICI-Expenses")
PDF_PASSWORD = os.getenv("ICICI_PDF_PASSWORD", "")

TRANSACTIONS_PATH = os.path.join(DATA_DIR, "icici_transactions.json")
PROCESSED_STMT_IDS_PATH = os.path.join(DATA_DIR, "processed_statement_ids.json")


def _get_service():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return build("gmail", "v1", credentials=creds)


def _get_processed_ids() -> set:
    if not os.path.exists(PROCESSED_STMT_IDS_PATH):
        return set()
    with open(PROCESSED_STMT_IDS_PATH) as f:
        return set(json.load(f))


def _save_processed_id(msg_id: str):
    os.makedirs(DATA_DIR, exist_ok=True)
    ids = _get_processed_ids()
    ids.add(msg_id)
    with open(PROCESSED_STMT_IDS_PATH, "w") as f:
        json.dump(list(ids), f)


def _load_transactions() -> list:
    if not os.path.exists(TRANSACTIONS_PATH):
        return []
    with open(TRANSACTIONS_PATH) as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def _save_transactions(transactions: list):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(TRANSACTIONS_PATH, "w") as f:
        json.dump(transactions, f, indent=2)


def _parse_pdf_transactions(pdf_bytes: bytes) -> list:
    """Extract transactions from ICICI statement PDF."""
    transactions = []

    try:
        pdf_file = io.BytesIO(pdf_bytes)
        with pdfplumber.open(pdf_file, password=PDF_PASSWORD) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text() or ""
                full_text += text + "\n"

                # Also try extracting tables
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row:
                            continue
                        row_text = " | ".join(str(c or "").strip() for c in row)
                        txn = _parse_transaction_row(row)
                        if txn:
                            transactions.append(txn)

            # If table extraction yielded nothing, try text parsing
            if not transactions:
                transactions = _parse_from_text(full_text)

    except Exception as e:
        print(f"PDF parse error: {e}")

    return transactions


def _parse_transaction_row(row: list) -> Optional[dict]:
    """Try to parse a table row as a transaction."""
    if not row or len(row) < 3:
        return None

    row_clean = [str(c or "").strip() for c in row]

    # Look for a date pattern in first column
    date_pattern = re.compile(r"\d{2}[/-]\d{2}[/-]\d{2,4}")
    if not date_pattern.search(row_clean[0]):
        return None

    # Try to find an amount (last non-empty numeric column)
    amount = None
    for cell in reversed(row_clean):
        cell = cell.replace(",", "").replace("Dr", "").replace("Cr", "").strip()
        try:
            amount = float(cell)
            break
        except ValueError:
            continue

    if amount is None or amount == 0:
        return None

    # Determine debit/credit
    row_str = " ".join(row_clean)
    txn_type = "credit" if "Cr" in row_str else "debit"

    # Description is usually middle columns
    description = " ".join(row_clean[1:-2]).strip() if len(row_clean) > 3 else row_clean[1] if len(row_clean) > 1 else ""
    description = re.sub(r"\s+", " ", description)

    return {
        "date": row_clean[0],
        "description": description,
        "amount": amount,
        "type": txn_type,
        "source": "icici_statement",
    }


def _parse_from_text(text: str) -> list:
    """Fallback: parse transactions from raw PDF text."""
    transactions = []

    # Common ICICI statement line pattern:
    # DD/MM/YYYY  MERCHANT NAME  CITY  DR/CR  AMOUNT
    pattern = re.compile(
        r"(\d{2}[/-]\d{2}[/-]\d{2,4})\s+"   # date
        r"(.+?)\s+"                            # description
        r"([\d,]+\.\d{2})\s*"                 # amount
        r"(Dr|Cr)?",                           # dr/cr
        re.IGNORECASE
    )

    for match in pattern.finditer(text):
        date_str, description, amount_str, dr_cr = match.groups()
        try:
            amount = float(amount_str.replace(",", ""))
        except ValueError:
            continue

        if amount < 1:
            continue

        transactions.append({
            "date": date_str.strip(),
            "description": description.strip(),
            "amount": amount,
            "type": "credit" if (dr_cr or "").lower() == "cr" else "debit",
            "source": "icici_statement",
        })

    return transactions


def _get_pdf_attachments(service, msg_id: str) -> list[bytes]:
    """Download all PDF attachments from a Gmail message."""
    pdfs = []
    msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
    parts = msg.get("payload", {}).get("parts", [])

    def walk_parts(parts):
        for part in parts:
            mime = part.get("mimeType", "")
            filename = part.get("filename", "")
            if mime == "application/pdf" or filename.lower().endswith(".pdf"):
                body = part.get("body", {})
                attachment_id = body.get("attachmentId")
                if attachment_id:
                    att = service.users().messages().attachments().get(
                        userId="me", messageId=msg_id, id=attachment_id
                    ).execute()
                    data = att.get("data", "")
                    if data:
                        pdfs.append(base64.urlsafe_b64decode(data))
                elif body.get("data"):
                    pdfs.append(base64.urlsafe_b64decode(body["data"]))
            if part.get("parts"):
                walk_parts(part["parts"])

    walk_parts(parts)
    return pdfs


def fetch_and_parse_statements(force_reprocess: bool = False) -> dict:
    """
    Fetch ICICI statement PDFs from Gmail, parse all transactions,
    merge with existing data. Returns summary dict.
    """
    service = _get_service()
    processed_ids = _get_processed_ids() if not force_reprocess else set()
    existing_transactions = _load_transactions()
    new_count = 0
    statements_processed = 0

    # Find ICICI-Expenses label
    labels_result = service.users().labels().list(userId="me").execute()
    label_id = next(
        (l["id"] for l in labels_result.get("labels", []) if l["name"] == ICICI_LABEL),
        None
    )
    if not label_id:
        print(f"Gmail label '{ICICI_LABEL}' not found.")
        return {"statements": 0, "transactions": 0, "new": 0}

    messages_result = service.users().messages().list(
        userId="me", labelIds=[label_id], maxResults=50
    ).execute()

    for msg_ref in messages_result.get("messages", []):
        msg_id = msg_ref["id"]
        if msg_id in processed_ids:
            continue

        pdfs = _get_pdf_attachments(service, msg_id)
        if not pdfs:
            continue

        for pdf_bytes in pdfs:
            txns = _parse_pdf_transactions(pdf_bytes)
            if txns:
                existing_transactions.extend(txns)
                new_count += len(txns)
                statements_processed += 1

        _save_processed_id(msg_id)

    _save_transactions(existing_transactions)
    print(f"Processed {statements_processed} statement(s), {new_count} new transaction(s).")
    return {
        "statements": statements_processed,
        "transactions": len(existing_transactions),
        "new": new_count,
    }


if __name__ == "__main__":
    result = fetch_and_parse_statements()
    print(result)
