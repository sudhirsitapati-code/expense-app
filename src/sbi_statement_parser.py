"""
sbi_statement_parser.py
Fetches SBI statement PDFs from Gmail, decrypts them, parses transactions,
and saves to icici_transactions (shared store) with account prefix SBI-*.
"""

import base64
import gc
import hashlib
import io
import os
import re
from datetime import datetime
from typing import Optional

import pdfplumber
from googleapiclient.discovery import build
from src.gmail_utils import get_credentials
from src import db as _db

SBI_LABEL = os.getenv("SBI_GMAIL_LABEL", "SBI-Expenses")
_SBI_PDF_PASSWORD = os.getenv("SBI_PDF_PASSWORD", "")
_SBI_PDF_PASSWORDS = [_SBI_PDF_PASSWORD] if _SBI_PDF_PASSWORD else []

FY27_START = datetime(2026, 4, 1)

_FY_MONTH_NO   = {4:1,5:2,6:3,7:4,8:5,9:6,10:7,11:8,12:9,1:10,2:11,3:12}
_FY_MONTH_NAME = {1:"Apr",2:"May",3:"Jun",4:"Jul",5:"Aug",6:"Sep",
                  7:"Oct",8:"Nov",9:"Dec",10:"Jan",11:"Feb",12:"Mar"}


def _get_service():
    return build("gmail", "v1", credentials=get_credentials())


def _get_processed_ids() -> set:
    return set(_db.load("processed_statement_ids", default=[]))


def _save_processed_id(msg_id: str):
    ids = _get_processed_ids()
    ids.add(msg_id)
    _db.save("processed_statement_ids", list(ids))


def _load_transactions() -> list:
    return _db.load("icici_transactions") or []


def _save_transactions(transactions: list):
    _db.save("icici_transactions", transactions)


def _make_txn_id(date: str, description: str, amount: float, direction: str = "") -> str:
    raw = f"sbi|{date}|{description}|{amount:.2f}|{direction}"
    return hashlib.sha1(raw.encode()).hexdigest()[:12]


def _parse_date(date_str: str) -> Optional[datetime]:
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%d.%m.%Y", "%d %b %Y", "%d %b %y"):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None


def _norm_date(date_str: str) -> str:
    dt = _parse_date(date_str)
    return dt.strftime("%d/%m/%Y") if dt else date_str


def _fy_fields(dt: Optional[datetime]) -> dict:
    if not dt:
        return {"month_no": None, "month_name": None}
    fn = _FY_MONTH_NO.get(dt.month, dt.month)
    return {"month_no": fn, "month_name": _FY_MONTH_NAME.get(fn, "")}


def _open_pdf(pdf_bytes: bytes):
    passwords = _SBI_PDF_PASSWORDS + [""]
    for pw in passwords:
        try:
            pdf = pdfplumber.open(io.BytesIO(pdf_bytes), password=pw)
            _ = pdf.pages
            return pdf
        except Exception:
            pass
    raise ValueError("No password worked for SBI PDF")


def _extract_account_from_text(text: str) -> str:
    """Extract last 4 digits of SBI account number."""
    for pat in [
        r"[Aa]ccount\s*[Nn]o\.?\s*:?\s*[\dX]{6,}(\d{4})\b",
        r"A/C\s*No\.?\s*:?\s*[\dX]{6,}(\d{4})\b",
        r"[Ss]avings\s+[Aa]ccount\s+[\dX]{6,}(\d{4})\b",
        r"\b(\d{11,})\b",  # SBI account numbers are typically 11 digits
    ]:
        m = re.search(pat, text[:3000])
        if m:
            digits = m.group(1) if len(m.groups()) >= 1 else m.group(0)
            return f"SBI-{digits[-4:]}"
    return "SBI-????"


def _parse_sbi_statement_text(text: str) -> list:
    """
    Parse SBI savings account e-statement.
    Format varies but common patterns:
      DD Mon YYYY  DESCRIPTION  DEBIT  CREDIT  BALANCE
      or
      DD/MM/YYYY  DESCRIPTION  DEBIT  CREDIT  BALANCE
    """
    transactions = []
    seen = set()

    # Pattern 1: DD Mon YYYY or DD/MM/YYYY followed by description and amounts
    # SBI statements often have: Date Particulars Debit Credit Balance
    pat = re.compile(
        r"(\d{2}[\/\-\s][A-Za-z]{3}[\/\-\s]\d{2,4}|\d{2}[\/\-]\d{2}[\/\-]\d{2,4})"
        r"\s+(.+?)\s+"
        r"([\d,]+\.\d{2})\s+"   # debit or first amount
        r"([\d,]+\.\d{2})\s+"   # credit or second amount
        r"(-?[\d,]+\.\d{2})",   # balance
        re.MULTILINE
    )
    prev_balance = None
    for m in pat.finditer(text):
        date_raw, desc, amt1_str, amt2_str, bal_str = m.groups()
        dt = _parse_date(date_raw.strip())
        if not dt:
            continue
        desc = desc.strip()
        if desc.upper() in ("PARTICULARS", "DATE", "DESCRIPTION", "TOTAL"):
            continue
        a1 = float(amt1_str.replace(",", ""))
        a2 = float(amt2_str.replace(",", ""))
        bal = float(bal_str.replace(",", ""))

        # Determine which is debit vs credit by balance delta
        if prev_balance is not None:
            delta = bal - prev_balance
            if abs(delta + a1) < 1:        # a1 is debit
                debit, credit = a1, 0.0
            elif abs(delta - a2) < 1:      # a2 is credit
                debit, credit = 0.0, a2
            elif a1 > 0 and a2 == 0:
                debit = a1; credit = 0.0
            elif a2 > 0 and a1 == 0:
                debit = 0.0; credit = a2
            else:
                debit = a1 if delta < 0 else 0.0
                credit = a2 if delta > 0 else 0.0
        else:
            if a1 > 0 and a2 == 0:
                debit, credit = a1, 0.0
            elif a2 > 0 and a1 == 0:
                debit, credit = 0.0, a2
            else:
                debit, credit = a1, 0.0

        prev_balance = bal
        amount = debit if debit > 0 else credit
        direction = "debit" if debit > 0 else "credit"
        if amount < 1:
            continue

        date_str = dt.strftime("%d/%m/%Y")
        key = f"{date_str}|{desc[:40]}|{amount}"
        if key in seen:
            continue
        seen.add(key)
        transactions.append({
            "date": date_str,
            "description": desc,
            "amount": amount,
            "txn_direction": direction,
            "balance": bal,
        })

    if transactions:
        return transactions

    # Pattern 2: single-amount format with Dr/Cr marker
    pat2 = re.compile(
        r"(\d{2}[\/\-\s][A-Za-z]{3}[\/\-\s]\d{2,4}|\d{2}[\/\-]\d{2}[\/\-]\d{2,4})"
        r"\s+(.+?)\s+"
        r"([\d,]+\.\d{2})\s*"
        r"(Dr\.?|Cr\.?|DR\.?|CR\.?)\b",
        re.MULTILINE | re.IGNORECASE
    )
    for m in pat2.finditer(text):
        date_raw, desc, amount_str, dr_cr = m.groups()
        dt = _parse_date(date_raw.strip())
        if not dt:
            continue
        amount = float(amount_str.replace(",", ""))
        if amount < 1:
            continue
        direction = "credit" if dr_cr.lower().startswith("cr") else "debit"
        date_str = dt.strftime("%d/%m/%Y")
        key = f"{date_str}|{desc[:40]}|{amount}"
        if key in seen:
            continue
        seen.add(key)
        transactions.append({
            "date": date_str,
            "description": desc.strip(),
            "amount": amount,
            "txn_direction": direction,
        })

    return transactions


def _parse_pdf_transactions(pdf_bytes: bytes) -> list:
    transactions = []
    account = "SBI-????"

    try:
        with _open_pdf(pdf_bytes) as pdf:
            full_text = ""
            for page in pdf.pages:
                full_text += (page.extract_text() or "") + "\n"

        account = _extract_account_from_text(full_text)
        transactions = _parse_sbi_statement_text(full_text)
        print(f"[sbi_parser] acct={account} txns={len(transactions)}")
        print(f"[sbi_parser] text_sample={repr(full_text[:300])}")
        for i, t in enumerate(transactions[:3]):
            print(f"[sbi_parser] txn[{i}] {t}")

    except Exception as e:
        print(f"[sbi_parser] PDF parse error: {e}")

    # Filter FY27 only
    fy27 = []
    for t in transactions:
        dt = _parse_date(t.get("date", ""))
        if dt and dt >= FY27_START:
            fy27.append(t)
    transactions = fy27

    # Enrich
    enriched = []
    for t in transactions:
        dt = _parse_date(t.get("date", ""))
        fy = _fy_fields(dt)
        debit  = t["amount"] if t.get("txn_direction") == "debit" else 0
        credit = t["amount"] if t.get("txn_direction") == "credit" else 0
        txn_id = _make_txn_id(t.get("date",""), t.get("description",""), t["amount"], t.get("txn_direction",""))
        enriched.append({
            "txn_id":             txn_id,
            "month_no":           fy["month_no"],
            "month_name":         fy["month_name"],
            "account":            account,
            "date":               t.get("date", ""),
            "transaction_details": t.get("description", ""),
            "paid_to":            None,
            "debit":              debit,
            "credit":             credit,
            "acc_type":           None,
            "heading":            None,
            "remarks":            "",
            "confidence":         None,
            "source":             "sbi_statement",
            "amount":             t["amount"],
            "type":               t.get("txn_direction", "debit"),
            "balance":            t.get("balance"),
            "bank":               "SBI",
        })

    return enriched


def _get_pdf_attachments(service, msg_id: str) -> list:
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


def fetch_and_parse_sbi_statements(force_reprocess: bool = False) -> dict:
    """
    Fetch SBI statement PDFs from Gmail, parse transactions,
    merge into shared icici_transactions store.
    """
    service = _get_service()
    processed_ids = _get_processed_ids() if not force_reprocess else set()
    existing = _load_transactions()
    existing_ids = {t.get("txn_id") for t in existing if t.get("txn_id")}
    new_count = 0
    statements_processed = 0

    label_query = f"label:{SBI_LABEL.lower().replace(' ', '-')}"
    messages_result = service.users().messages().list(
        userId="me",
        q=f"(from:alerts@sbi.co.in OR from:sbialerts@sbi.co.in OR from:noreply@onlinesbi.com "
          f"OR from:yonobysbi@alerts.sbi.bank.in OR {label_query}) "
          "has:attachment filename:pdf",
        maxResults=10
    ).execute()

    msgs_to_process = [
        m for m in messages_result.get("messages", [])
        if m["id"] not in processed_ids
    ]
    print(f"[sbi] Found {len(msgs_to_process)} unprocessed SBI emails")

    for msg_ref in msgs_to_process:
        msg_id = msg_ref["id"]
        try:
            pdfs = _get_pdf_attachments(service, msg_id)
        except Exception as e:
            print(f"[sbi] Failed to fetch attachments for {msg_id}: {e}")
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
                print(f"[sbi] Parse error in message {msg_id}: {e}")
            finally:
                del pdf_bytes
                gc.collect()

        _save_processed_id(msg_id)
        _save_transactions(existing)

    print(f"[sbi] Processed {statements_processed} statement(s), {new_count} new transaction(s).")
    return {
        "statements": statements_processed,
        "transactions": len(existing),
        "new": new_count,
    }
