"""
icici_statement_parser.py
Fetches ICICI statement PDFs from Gmail, decrypts them, parses all transactions
with full ACC26-compatible fields, and saves to data/icici_transactions.json.
"""

import base64
import hashlib
import io
import json
import os
import re
from datetime import datetime
from typing import Optional

import pdfplumber
from googleapiclient.discovery import build
from src.gmail_utils import get_credentials
from src.icici_classifier import classify_transactions

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
ICICI_LABEL = os.getenv("ICICI_GMAIL_LABEL", "ICICI-Expenses")
PDF_PASSWORD = os.getenv("ICICI_PDF_PASSWORD", "")

TRANSACTIONS_PATH = os.path.join(DATA_DIR, "icici_transactions.json")
PROCESSED_STMT_IDS_PATH = os.path.join(DATA_DIR, "processed_statement_ids.json")

# FY month mapping: calendar month → FY month number (Apr=1 … Mar=12)
_FY_MONTH_NO = {4:1,5:2,6:3,7:4,8:5,9:6,10:7,11:8,12:9,1:10,2:11,3:12}
_FY_MONTH_NAME = {1:"Apr",2:"May",3:"Jun",4:"Jul",5:"Aug",6:"Sep",
                  7:"Oct",8:"Nov",9:"Dec",10:"Jan",11:"Feb",12:"Mar"}


def _get_service():
    return build("gmail", "v1", credentials=get_credentials())


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


def _make_txn_id(date: str, description: str, amount: float) -> str:
    raw = f"{date}|{description}|{amount:.2f}"
    return hashlib.sha1(raw.encode()).hexdigest()[:12]


def _parse_date(date_str: str) -> Optional[datetime]:
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None


def _fy_fields(dt: Optional[datetime]) -> dict:
    if not dt:
        return {"month_no": None, "month_name": None}
    fn = _FY_MONTH_NO.get(dt.month, dt.month)
    return {"month_no": fn, "month_name": _FY_MONTH_NAME.get(fn, "")}


def _extract_account_from_text(text: str) -> str:
    """Try to extract the last 4 digits of the ICICI account from statement header."""
    m = re.search(r"[Aa]ccount\s*[Nn]o[\.\:]?\s*[\dX]{4,}(\d{4})", text)
    if m:
        return f"icic{m.group(1)}"
    m = re.search(r"(\d{4})\s*$", text[:500], re.MULTILINE)
    if m:
        return f"icic{m.group(1)}"
    return "icici"


def _parse_pdf_transactions(pdf_bytes: bytes) -> list:
    """Extract and classify transactions from ICICI statement PDF."""
    transactions = []
    account = "icici"

    try:
        pdf_file = io.BytesIO(pdf_bytes)
        with pdfplumber.open(pdf_file, password=PDF_PASSWORD) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text() or ""
                full_text += text + "\n"

                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row:
                            continue
                        txn = _parse_transaction_row(row)
                        if txn:
                            transactions.append(txn)

            # Try to get account from header
            account = _extract_account_from_text(full_text)

            # Fallback to text parsing if no tables worked
            if not transactions:
                transactions = _parse_from_text(full_text)

    except Exception as e:
        print(f"PDF parse error: {e}")

    # Enrich with ACC26 fields
    enriched = []
    for t in transactions:
        dt = _parse_date(t.get("date", ""))
        fy = _fy_fields(dt)
        debit = t["amount"] if t.get("txn_direction") == "debit" else 0
        credit = t["amount"] if t.get("txn_direction") == "credit" else 0
        txn_id = _make_txn_id(t.get("date",""), t.get("description",""), t["amount"])
        enriched.append({
            "txn_id": txn_id,
            "month_no": fy["month_no"],
            "month_name": fy["month_name"],
            "account": account,
            "date": t.get("date", ""),
            "transaction_details": t.get("description", ""),
            "paid_to": None,
            "debit": debit,
            "credit": credit,
            "acc_type": None,
            "heading": None,
            "remarks": "",
            "confidence": None,
            "source": "icici_statement",
            # legacy field kept for reconciliation
            "amount": t["amount"],
            "type": t.get("txn_direction", "debit"),
        })

    # AI + rule-based classification
    if enriched:
        for t in enriched:
            t["description"] = t["transaction_details"]
        classify_transactions(enriched)
        for t in enriched:
            t.pop("description", None)

    return enriched


def _parse_transaction_row(row: list) -> Optional[dict]:
    """Try to parse a table row as a transaction."""
    if not row or len(row) < 3:
        return None

    row_clean = [str(c or "").strip() for c in row]

    date_pattern = re.compile(r"\d{2}[/-]\d{2}[/-]\d{2,4}")
    if not date_pattern.search(row_clean[0]):
        return None

    amount = None
    for cell in reversed(row_clean):
        cell_clean = cell.replace(",", "").replace("Dr", "").replace("Cr", "").strip()
        try:
            val = float(cell_clean)
            if val > 0:
                amount = val
                break
        except ValueError:
            continue

    if amount is None:
        return None

    row_str = " ".join(row_clean)
    direction = "credit" if "Cr" in row_str else "debit"
    description = " ".join(row_clean[1:-2]).strip() if len(row_clean) > 3 else (row_clean[1] if len(row_clean) > 1 else "")
    description = re.sub(r"\s+", " ", description)

    return {"date": row_clean[0], "description": description, "amount": amount, "txn_direction": direction}


def _parse_from_text(text: str) -> list:
    """Fallback: parse transactions from raw PDF text."""
    transactions = []
    pattern = re.compile(
        r"(\d{2}[/-]\d{2}[/-]\d{2,4})\s+(.+?)\s+([\d,]+\.\d{2})\s*(Dr|Cr)?",
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
            "txn_direction": "credit" if (dr_cr or "").lower() == "cr" else "debit",
        })
    return transactions


def _get_pdf_attachments(service, msg_id: str) -> list:
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
    Fetch ICICI statement PDFs from Gmail, parse all transactions with ACC26 fields,
    merge (deduplicate by txn_id) with existing data. Returns summary dict.
    """
    service = _get_service()
    processed_ids = _get_processed_ids() if not force_reprocess else set()
    existing = _load_transactions()
    existing_ids = {t.get("txn_id") for t in existing if t.get("txn_id")}
    new_count = 0
    statements_processed = 0

    labels_result = service.users().labels().list(userId="me").execute()
    label_id = next(
        (l["id"] for l in labels_result.get("labels", []) if l["name"] == ICICI_LABEL),
        None
    )
    if not label_id:
        print(f"Gmail label '{ICICI_LABEL}' not found.")
        return {"statements": 0, "transactions": len(existing), "new": 0}

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
            added = 0
            for t in txns:
                tid = t.get("txn_id")
                if tid and tid not in existing_ids:
                    existing.append(t)
                    existing_ids.add(tid)
                    added += 1
            if added:
                new_count += added
                statements_processed += 1

        _save_processed_id(msg_id)

    _save_transactions(existing)
    print(f"Processed {statements_processed} statement(s), {new_count} new transaction(s).")
    return {
        "statements": statements_processed,
        "transactions": len(existing),
        "new": new_count,
    }


if __name__ == "__main__":
    result = fetch_and_parse_statements()
    print(result)
