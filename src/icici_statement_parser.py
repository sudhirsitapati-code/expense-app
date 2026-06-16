"""
icici_statement_parser.py
Fetches ICICI statement PDFs from Gmail, decrypts them, parses all transactions
with full ACC26-compatible fields, and saves to data/icici_transactions.json.
"""

import base64
import gc
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
# Try both cases — summary statements use SUDH3108, e-statements use sudh3108
_PDF_PASSWORD_PRIMARY = os.getenv("ICICI_PDF_PASSWORD", "SUDH3108")
_PDF_PASSWORDS = [_PDF_PASSWORD_PRIMARY, _PDF_PASSWORD_PRIMARY.lower(), _PDF_PASSWORD_PRIMARY.upper()]
PDF_PASSWORD = _PDF_PASSWORD_PRIMARY  # kept for legacy callers

TRANSACTIONS_PATH = os.path.join(DATA_DIR, "icici_transactions.json")
PROCESSED_STMT_IDS_PATH = os.path.join(DATA_DIR, "processed_statement_ids.json")

# FY month mapping: calendar month → FY month number (Apr=1 … Mar=12)
_FY_MONTH_NO = {4:1,5:2,6:3,7:4,8:5,9:6,10:7,11:8,12:9,1:10,2:11,3:12}
_FY_MONTH_NAME = {1:"Apr",2:"May",3:"Jun",4:"Jul",5:"Aug",6:"Sep",
                  7:"Oct",8:"Nov",9:"Dec",10:"Jan",11:"Feb",12:"Mar"}


def _get_service():
    return build("gmail", "v1", credentials=get_credentials())


from src import db as _db


def _get_processed_ids() -> set:
    return set(_db.load("processed_statement_ids", default=[]))


def _save_processed_id(msg_id: str):
    ids = _get_processed_ids()
    ids.add(msg_id)
    _db.save("processed_statement_ids", list(ids))


def _load_transactions() -> list:
    return _db.load("icici_transactions")


def _save_transactions(transactions: list):
    _db.save("icici_transactions", transactions)


def _make_txn_id(date: str, description: str, amount: float, direction: str = "") -> str:
    raw = f"{date}|{description}|{amount:.2f}|{direction}"
    return hashlib.sha1(raw.encode()).hexdigest()[:12]


def _parse_date(date_str: str) -> Optional[datetime]:
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%d.%m.%Y"):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None


def _norm_date(date_str: str) -> str:
    """Normalise any supported date format to DD/MM/YYYY."""
    dt = _parse_date(date_str)
    return dt.strftime("%d/%m/%Y") if dt else date_str


def _fy_fields(dt: Optional[datetime]) -> dict:
    if not dt:
        return {"month_no": None, "month_name": None}
    fn = _FY_MONTH_NO.get(dt.month, dt.month)
    return {"month_no": fn, "month_name": _FY_MONTH_NAME.get(fn, "")}



def _open_pdf(pdf_bytes: bytes):
    """Try all password variants; return open pdfplumber object or raise."""
    for pw in _PDF_PASSWORDS:
        try:
            pdf = pdfplumber.open(io.BytesIO(pdf_bytes), password=pw)
            _ = pdf.pages  # force open
            return pdf
        except Exception:
            pass
    raise ValueError("No password variant worked for this PDF")


FY27_START = datetime(2026, 4, 1)


def _parse_pdf_transactions(pdf_bytes: bytes) -> list:
    """Extract and classify transactions from ICICI statement PDF."""
    transactions = []
    account = "icici"

    try:
        with _open_pdf(pdf_bytes) as pdf:
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

            # Detect savings vs OD format and use the right text parser
            savings_txns = _parse_savings_statement_text(full_text)
            od_txns = _parse_from_text(full_text)
            best_text = savings_txns if len(savings_txns) >= len(od_txns) else od_txns
            if best_text and len(best_text) > len(transactions):
                transactions = best_text

    except Exception as e:
        print(f"PDF parse error: {e}")

    # Filter to FY27 only (April 1, 2026 onwards)
    fy27_transactions = []
    for t in transactions:
        dt = _parse_date(t.get("date", ""))
        if dt and dt < FY27_START:
            continue
        fy27_transactions.append(t)
    transactions = fy27_transactions

    # Enrich with ACC26 fields
    enriched = []
    for t in transactions:
        dt = _parse_date(t.get("date", ""))
        fy = _fy_fields(dt)
        debit = t["amount"] if t.get("txn_direction") == "debit" else 0
        credit = t["amount"] if t.get("txn_direction") == "credit" else 0
        txn_id = _make_txn_id(t.get("date",""), t.get("description",""), t["amount"], t.get("txn_direction",""))
        enriched.append({
            "txn_id": txn_id,
            "month_no": fy["month_no"],
            "month_name": fy["month_name"],
            "account": account,
            "date": t.get("date", ""),
            "transaction_details": t.get("description", ""),
            # paid_to_hint from savings parser = short bold name; used as fallback if classifier can't identify payee
            "paid_to": t.get("paid_to_hint") or None,
            "debit": debit,
            "credit": credit,
            "acc_type": None,
            "heading": None,
            "remarks": "",
            "confidence": None,
            "source": "icici_statement",
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
            # Keep paid_to from classifier if set; fall back to hint
            if not t.get("paid_to") and t.get("paid_to_hint"):
                t["paid_to"] = t["paid_to_hint"]
            t.pop("paid_to_hint", None)

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


def _extract_account_from_text(text: str) -> str:
    """Extract last 4 digits of account number from statement text."""
    # Prefer masked account number pattern: XXXXXXXX7281 (8+ X's followed by digits)
    # over Customer ID (XXXXX6511 = shorter X sequence)
    for pat in [
        r"[Xx]{6,}(\d{4})\b",                                      # long masked account e.g. XXXXXXXX7281
        r"[Aa]ccount\s*(?:[Nn]o|[Nn]umber)[\.\:]?\s*[\dX]{6,}(\d{4})",  # Account No: XXXXXXXX7281
        r"[Aa]/[Cc]\s+[Xx]{4,}(\d{4})\b",                         # A/c XXXX1234
        r"Statement.*?[Aa]ccount\s+[Xx\d]{6,}(\d{4})",            # in statement header line
    ]:
        m = re.search(pat, text[:2000])
        if m:
            return f"ICICI-{m.group(1)}"
    # Fallback: full numeric account number (10+ digits like 003801011331)
    m = re.search(r"\b\d{6,}(\d{4})\b", text[:3000])
    if m:
        return f"ICICI-{m.group(1)}"
    return "ICICI"


def _parse_savings_statement_text(text: str) -> list:
    """
    Parse ICICI savings account statement (e.g. account 1331) which has format:
      PAID_TO_SHORT_NAME           ← bold first line (payee)
      S_NO DD.MM.YYYY AMOUNT BAL   ← sequence, dot-date, single amount, balance
      FULL/REFERENCE/DETAILS...    ← UPI/NEFT/BIL reference (may be multi-line)

    Uses balance delta to determine debit vs credit.
    """
    lines = text.split("\n")

    # Row line: sequence_number  DD.MM.YYYY  amount  balance
    row_pat = re.compile(
        r"^(\d{1,4})\s+(\d{2}\.\d{2}\.\d{4})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$"
    )

    row_indices = [i for i, ln in enumerate(lines) if row_pat.match(ln.strip())]
    if not row_indices:
        return []

    transactions = []
    prev_balance = None
    seen = set()

    for idx, row_i in enumerate(row_indices):
        m = row_pat.match(lines[row_i].strip())
        _, date_dot, amount_str, balance_str = m.groups()

        # Convert DD.MM.YYYY → DD/MM/YYYY (canonical date format)
        d, mo, yr = date_dot.split(".")
        date_str = f"{d}/{mo}/{yr}"

        # paid_to: closest non-empty, non-header line above the row line
        paid_to = ""
        for j in range(row_i - 1, max(row_i - 4, -1), -1):
            candidate = lines[j].strip()
            if candidate and not row_pat.match(candidate):
                # Skip column header words
                if candidate.lower() not in ("date", "balance", "amount (inr)", "withdrawal amount (inr)", "deposit amount (inr)"):
                    paid_to = candidate
                    break

        # description: lines between this row and next paid_to
        next_row_i = row_indices[idx + 1] if idx + 1 < len(row_indices) else len(lines)
        # The line at next_row_i - 1 is the paid_to of the next txn — exclude it
        desc_end = next_row_i - 1 if idx + 1 < len(row_indices) else len(lines)
        desc_lines = [lines[j].strip() for j in range(row_i + 1, desc_end) if lines[j].strip()]
        full_desc = " ".join(desc_lines)

        amount = float(amount_str.replace(",", ""))
        balance = float(balance_str.replace(",", ""))

        if amount < 1:
            prev_balance = balance
            continue

        direction = "credit" if (prev_balance is not None and balance > prev_balance) else "debit"
        prev_balance = balance

        # Use full_desc as description; paid_to as short payee
        description = full_desc or paid_to

        # Override: these transaction types are always debits regardless of balance delta
        ALWAYS_DEBIT_PREFIXES = ("smp/", "bil/", "emi/", "nach/", "mandate/")
        if any(description.lower().startswith(p) or paid_to.lower().startswith(p) for p in ALWAYS_DEBIT_PREFIXES):
            direction = "debit"

        key = f"{date_str}|{description[:40]}|{amount}|{balance_str}"
        if key in seen:
            continue
        seen.add(key)

        transactions.append({
            "date": date_str,
            "description": description,
            "paid_to_hint": paid_to,   # passed through for enrichment
            "amount": amount,
            "txn_direction": direction,
        })

    return transactions


def _parse_from_text(text: str) -> list:
    """
    Parse transactions from raw ICICI statement text.
    Handles two column formats:
      - Savings/legacy: ...amount Dr/Cr
      - Current/OD e-statement: date particulars deposits withdrawals balance
    Uses running balance to determine debit/credit direction for single-amount rows.
    """
    transactions = []

    # Format 1: ICICI e-statement (current/savings detailed)
    # Line: DD-MM-YYYY  PARTICULARS  [DEPOSITS]  [WITHDRAWALS]  BALANCE
    # Note: no MODE token captured — MODE is often absent or part of PARTICULARS
    efmt = re.compile(
        r"^(\d{2}-\d{2}-\d{4})\s+"        # date
        r"(.+?)\s+"                         # particulars (lazy, stops at first amount)
        r"([\d,]+\.\d{2})\s+"              # amount1
        r"(?:([\d,]+\.\d{2})\s+)?"         # amount2 optional
        r"(-?[\d,]+\.\d{2})\s*$",          # balance (may be negative for OD accounts)
        re.MULTILINE
    )

    # Extract opening B/F balance so we can track direction from balance changes
    bf_match = re.search(r"\d{2}-\d{2}-\d{4}\s+B/F\s+(-?[\d,]+\.\d{2})", text, re.IGNORECASE)
    prev_balance = float(bf_match.group(1).replace(",", "")) if bf_match else None

    seen = set()
    for m in efmt.finditer(text):
        date_str, desc, amt1, amt2, balance_str = m.groups()
        desc = desc.strip()
        # Skip header / B/F / Total rows
        if desc.upper() in ("PARTICULARS", "MODE PARTICULARS", "DATE"):
            continue
        if "B/F" in desc.upper():
            prev_balance = float(balance_str.replace(",", ""))
            continue
        if desc.strip().rstrip(":").upper() == "TOTAL":
            continue

        balance = float(balance_str.replace(",", ""))
        a1 = float(amt1.replace(",", ""))
        a2 = float(amt2.replace(",", "")) if amt2 else 0.0

        if amt2:
            # Two amounts present: deposits column and withdrawals column
            # Pick the non-zero one; withdrawal (debit) takes priority
            if a2 > 0:
                amount, direction = a2, "debit"
            else:
                amount, direction = a1, "credit"
        else:
            amount = a1
            if prev_balance is not None:
                # Balance went up (or less negative) → credit; down (more negative) → debit
                direction = "credit" if balance > prev_balance else "debit"
            else:
                direction = "debit"  # safe default

        prev_balance = balance

        if amount < 1:
            continue
        # Include balance in key so same-amount debit+credit pair aren't deduped
        key = f"{date_str}|{desc}|{amount}|{balance_str}"
        if key in seen:
            continue
        seen.add(key)
        transactions.append({
            "date": _norm_date(date_str),
            "description": desc,
            "amount": amount,
            "txn_direction": direction,
        })

    if transactions:
        return transactions

    # Format 2: legacy Dr/Cr format
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
            "date": _norm_date(date_str.strip()),
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
    if force_reprocess:
        # Drop statement-imported transactions so re-parse replaces them cleanly
        existing = [t for t in _load_transactions() if t.get("source") != "icici_statement"]
    else:
        existing = _load_transactions()
    existing_ids = {t.get("txn_id") for t in existing if t.get("txn_id")}
    new_count = 0
    statements_processed = 0

    # Search by ICICI sender OR the ICICI label (catches forwarded emails from personal Gmail)
    label_query = f"label:{ICICI_LABEL.lower().replace(' ', '-')}"
    messages_result = service.users().messages().list(
        userId="me",
        q=f"(from:customernotification@icici.bank.in OR from:alert@icici.bank.in "
          f"OR from:estatement@icici.bank.in OR {label_query}) "
          "has:attachment filename:pdf",
        maxResults=5
    ).execute()

    msgs_to_process = [
        m for m in messages_result.get("messages", [])
        if m["id"] not in processed_ids
    ]

    for msg_ref in msgs_to_process:
        msg_id = msg_ref["id"]
        try:
            pdfs = _get_pdf_attachments(service, msg_id)
        except Exception as e:
            print(f"[stmt] Failed to fetch attachments for {msg_id}: {e}")
            _save_processed_id(msg_id)
            continue

        for pdf_bytes in pdfs:
            try:
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
            except Exception as e:
                print(f"[stmt] Parse error in message {msg_id}: {e}")
            finally:
                del pdf_bytes
                gc.collect()

        _save_processed_id(msg_id)
        # Save incrementally so progress isn't lost if next email OOMs
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
