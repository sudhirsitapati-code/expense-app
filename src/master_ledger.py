"""
master_ledger.py
Unified transaction ledger for all 7 bank accounts.
Parses Gmail alerts, applies business rules, classifies with AI.

Accounts handled:
  - 4 ICICI savings accounts (detected from alert account number)
  - 1 SBI savings account
  - 2 credit cards (ICICI CC / SBI Card)

Business rules:
  - Cash basis: expense recorded when cash paid
  - Credit card: accrued on spend; tracked as Credit Card Loan (investment)
  - Cash to Shiloj / Mohammed: Short Term Advance (investment) until bill submitted
  - NEFT/IMPS/UPI between own accounts: Transfer
  - Loan EMI, SIP, MF: Investment
  - Salary, interest credits: Income
  - GCPL reimbursements: Official
"""

import hashlib
import json
import os
import re
from datetime import datetime, timedelta
from typing import Optional

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")

LEDGER_PATH        = os.path.join(DATA_DIR, "master_ledger.json")
CC_BALANCE_PATH    = os.path.join(DATA_DIR, "cc_balance.json")

# FY calendar mapping
CAL_TO_FY_NO   = {4:1,5:2,6:3,7:4,8:5,9:6,10:7,11:8,12:9,1:10,2:11,3:12}
CAL_TO_FY_MON  = {4:"Apr",5:"May",6:"Jun",7:"Jul",8:"Aug",9:"Sep",
                   10:"Oct",11:"Nov",12:"Dec",1:"Jan",2:"Feb",3:"Mar"}

# ── Business rule keywords ────────────────────────────────────────────────────

STAFF_NAMES = ["shiloj", "mohammed", "mohamad", "shiloj james", "md "]

TRANSFER_KEYWORDS = [
    "neft", "imps", "rtgs", "trf to", "transfer to", "fund transfer",
    "internal transfer", "self transfer", "upi/", "upi-",
    "@okaxis", "@oksbi", "@okicici", "@ybl", "@ibl",
    "bil/neft", "bil/imps", ":transfer ", "007281:", "011331:",
    "trfr to", "fund trf", "online transfer",
    "inft/", "/inft/", "net banking inf", "inf/inft", "/inf/",
]
INCOME_KEYWORDS = [
    "salary cr", "sal cr", "salary credit", "interest credit", "int cr",
    "dividend", "refund", "reversal cr", "credit reversal", "cashback",
    "gcpl salary", "godrej salary",
]
INVESTMENT_KEYWORDS = [
    "sip", "mutual fund", "mf/", "nps", "ppf", "emi", "loan emi",
    "infina finance", "mizugami", "home loan", "mortgage",
    "investment", "fd open", "rd open",
]
OFFICIAL_KEYWORDS = [
    "gcpl", "godrej consumer", "reimbursement", "reimb", "official",
]

# ── ACC26 heading classification rules (keyword → heading) ────────────────────

HEADING_RULES = [
    (["grocery", "bigbasket", "dmart", "reliance fresh", "nature basket",
      "grofer", "blinkit", "zepto", "swiggy instamart", "jiomart"], "Groceries"),
    (["salary", "sal ", "staff pay", "domestic", "cook pay",
      "driver pay", "maid", "nanny"], "Staff Salary"),
    (["electricity", "mahadiscom", "msedcl", "best ", "gas ", "igl ",
      "piped gas", "adani gas", "bses", "tata power", "torrent power"], "Electricity & Gas"),
    (["kalpataru", "maintenance soc", "society maint", "housing soc",
      "bmc ", "icard", "icic0018 soc"], "Kalpataru Maintenance"),
    (["school", "tuition", "coaching", "iit", "icse", "cbse",
      "symbiosis", "cambridge", "academy", "college fee",
      "children education", "ed fee"], "Children Education"),
    (["ketki", "ketaki", "wife", "mrs sitapati"], "Ketki"),
    (["hotel", "holiday", "airbnb", "makemytrip", "yatra", "cleartrip",
      "flight", "indigo", "air india", "vistara", "spicejet",
      "booking.com", "trivago", "oyo rooms", "travel"], "Holiday"),
    (["kalpalata", "charity", "donation", "trust ", "foundation",
      "ngo ", "relief fund"], "Charity"),
    (["uspaar", "saving ", "frugal"], "Uspaar"),
    (["malhar", "building ", "property tax", "construction",
      "architect", "interior", "renovation"], "Malhar"),
    (["amma", "mother", "mom pay", "ammal"], "Amma"),
    (["doctor", "hospital", "pharmacy", "medical", "apollo",
      "fortis", "lilavati", "hinduja", "nanawati", "medic",
      "clinic", "diagnostic", "pathology", "lab test"], "Medical"),
    (["gym", "fitness", "cult.fit", "yoga", "spa ", "bombay gymkhana",
      "gymkhana", "salon", "hair ", "massage", "wellness"], "Wellness"),
    (["zara", "h&m", "marks & spencer", "westside", "shoppers stop",
      "myntra", "ajio", "fabindia", "lifestyle store", "clothes",
      "garment", "fashion", "shirt", "trouser"], "Clothes"),
    (["amazon gift", "gift ", "present ", "birthday", "anniversary gift",
      "flower"], "Gifts"),
    (["swiggy", "zomato", "dineout", "eazydiner", "restaurant",
      "cafe ", "bistro ", "dining", "food delivery", "eatout"], "Eating Out"),
    (["netflix", "amazon prime", "hotstar", "spotify", "apple music",
      "pvr ", "inox ", "bookmyshow", "entertainment", "gaming",
      "playstation", "steam"], "Entertainment"),
    (["insurance", "lic ", "hdfc life", "max life", "icici pru",
      "bajaj allianz", "star health", "mediclai"], "Insurance"),
    (["home loan", "housing loan", "emi sbi", "emi icici",
      "loan repay", "mortgage emi"], "Home Loan"),
    (["income tax", "tds ", "gst ", "advance tax",
      "tax payment", "oltas"], "Tax"),
    (["atm ", "atm/", "cash withdrawal", "cash wd"], "Cash"),
    (["alcohol", "wine ", "beer ", "whisky", "vodka", "liquor",
      "beverage", "bar tab"], "Alcohol"),
    (["amazon", "flipkart", "meesho", "nykaa", "misc purchase",
      "online purchase"], "Misc"),
    (["maintenance", "plumber", "electrician", "carpenter", "repair",
      "servicing", "ac service", "pest control"], "Maintenance Expense"),
    (["office supply", "stationery", "printer", "laptop", "monitor",
      "home office", "wfh"], "Home office"),
    (["one time", "special purchase", "luxury", "appliance",
      "furniture", "artwork", "antique"], "One Time Charge"),
    (["financial charge", "bank charge", "od interest", "overdraft",
      "late fee", "penalty"], "Financial Expense / OD Interest"),
]


# ── Helpers ───────────────────────────────────────────────────────────────────

from src import db as _db

_PATH_TO_KEY = {
    LEDGER_PATH: "master_ledger",
}


def _load_json(path):
    key = _PATH_TO_KEY.get(path)
    if key:
        return _db.load(key)
    if not os.path.exists(path):
        return []
    with open(path) as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def _save_json(path, data):
    key = _PATH_TO_KEY.get(path)
    if key:
        _db.save(key, data)
        return
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(data, f, indent=2, default=str)


def _txn_id(date: str, account: str, description: str, amount: float) -> str:
    raw = f"{date}|{account}|{description}|{amount}"
    return hashlib.sha1(raw.encode()).hexdigest()[:16]


def _parse_date(s: str) -> Optional[datetime]:
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s.strip(), fmt)
        except ValueError:
            continue
    return None


def _fy_info(dt: datetime) -> dict:
    return {
        "fy_month_no":   CAL_TO_FY_NO[dt.month],
        "fy_month_name": CAL_TO_FY_MON[dt.month],
        "fy_year":       dt.year if dt.month >= 4 else dt.year - 1,
    }


# ── Rule-based classifier ─────────────────────────────────────────────────────

def _classify_type(desc: str, account_type: str, paid_to: str) -> tuple[str, bool]:
    """Returns (type, certain)."""
    d = desc.lower()
    p = paid_to.lower()

    # Staff cash advance
    if any(s in p or s in d for s in STAFF_NAMES):
        return "investment", True   # Short Term Advance until bill

    if account_type == "credit_card":
        return "expense", True      # Credit card: accrue on spend

    if any(k in d for k in INCOME_KEYWORDS):
        return "income", True

    if any(k in d for k in INVESTMENT_KEYWORDS):
        return "investment", True

    if any(k in d for k in OFFICIAL_KEYWORDS):
        return "official", True

    if any(k in d for k in TRANSFER_KEYWORDS):
        return "transfer", True

    return "expense", False         # Default; uncertain


def _classify_heading(desc: str, txn_type: str) -> tuple[Optional[str], bool]:
    """Returns (heading, certain)."""
    if txn_type == "transfer":
        return "Interbank", True
    if txn_type in ("income", "official"):
        return None, True

    if txn_type == "investment":
        d = desc.lower()
        if any(k in d for k in STAFF_NAMES):
            return "Short Term Advance", True
        if any(k in d for k in ["home loan", "mortgage", "housing loan"]):
            return "Home Loan", True
        if any(k in d for k in ["insurance", "lic", "premium"]):
            return "Insurance", True
        if any(k in d for k in ["sip", "mutual fund", "mf/"]):
            return "Investment", True
        return "Investment", False

    d = desc.lower()
    for keywords, heading in HEADING_RULES:
        if any(k in d for k in keywords):
            return heading, True

    return "Misc", False            # Uncertain default


def _extract_paid_to(desc: str, bank: str) -> str:
    """Clean up raw bank description to a readable vendor name."""
    cleaned = desc.strip()

    # Strip VPS* payment gateway prefix (ICICI debit card POS)
    cleaned = re.sub(r"^VPS\*", "", cleaned, flags=re.I).strip()
    # Strip other common prefixes
    for p in [r"^UPI-", r"^NEFT-", r"^IMPS-", r"^POS-", r"^ATM-",
              r"^UPI/", r"^NEFT/", r"^IMPS/", r"^POS ",
              r"^PURCHASE AT ", r"^BILL PAYMENT-", r"^ECS-",
              r"\d{6,}/", r"/\d{4,}/"]:
        cleaned = re.sub(p, "", cleaned, flags=re.I).strip()

    # Known VPS* → readable name mappings
    VENDOR_MAP = {
        "COPPER CHIM": "Copper Chimney", "COPPER CHIMNEY": "Copper Chimney",
        "HAMLEYS": "Hamleys", "MANSURI CAT": "Mansuri Caterers",
        "BOMBAY GYM": "Bombay Gymkhana", "BIGBASKET": "BigBasket",
        "SWIGGY": "Swiggy", "ZOMATO": "Zomato",
    }
    upper = cleaned.upper()
    for k, v in VENDOR_MAP.items():
        if k in upper:
            return v

    # Take first meaningful part
    parts = re.split(r"[/|\\]", cleaned)
    vendor = parts[0].strip()[:50].title()
    return vendor or desc[:30].title()


def classify_transaction(txn: dict) -> dict:
    """Apply rules + populate type, heading, paid_to, uncertain flag."""
    desc  = txn.get("raw_description", "")
    acct  = txn.get("account_type", "savings")
    existing_paid_to = txn.get("paid_to", "")

    paid_to = existing_paid_to or _extract_paid_to(desc, txn.get("bank", ""))
    txn["paid_to"] = paid_to

    # Only classify if not already manually set
    if txn.get("confidence") == "manual":
        return txn

    txn_type, type_certain = _classify_type(desc, acct, paid_to)
    heading, heading_certain = _classify_heading(desc, txn_type)

    txn["type"]    = txn_type
    txn["heading"] = heading

    uncertain_fields = []
    if not type_certain:
        uncertain_fields.append("type")
    if heading and not heading_certain:
        uncertain_fields.append("heading")

    txn["uncertain"]        = bool(uncertain_fields)
    txn["uncertain_fields"] = uncertain_fields
    txn["confidence"]       = "rule"

    return txn


# ── AI batch classifier (Azure OpenAI) ────────────────────────────────────────

def ai_classify_batch(transactions: list[dict]) -> list[dict]:
    """Use Azure OpenAI to classify uncertain transactions in batches of 15."""
    try:
        from openai import AzureOpenAI
        client = AzureOpenAI(
            api_key=os.getenv("AZURE_OPENAI_KEY"),
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
            api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-12-01-preview"),
        )
        deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-5.5")
    except Exception:
        return transactions

    BATCH = 15
    type_values    = ["expense","income","investment","transfer","official"]
    heading_values = [
        "Groceries","Staff Salary","Electricity & Gas","Misc","Cash","Alcohol",
        "Wellness","Clothes","Gifts","Medical","Amma","Ketki","Children Education",
        "Charity","Uspaar","Holiday","Eating Out","Entertainment","Malhar",
        "Maintenance Expense","Home office","One Time Charge","Kalpataru Maintenance",
        "Financial Expense / OD Interest","Insurance","Home Loan","Tax",
        "Short Term Advance","Credit Card Loan","Investment",
    ]

    uncertain = [t for t in transactions if t.get("uncertain") and t.get("confidence") != "manual"]
    if not uncertain:
        return transactions

    for i in range(0, len(uncertain), BATCH):
        batch = uncertain[i:i + BATCH]
        items = [
            {"id": t["txn_id"], "desc": t["raw_description"],
             "amount": t.get("debit") or t.get("credit",0),
             "bank": t.get("bank",""), "account_type": t.get("account_type","savings")}
            for t in batch
        ]
        prompt = f"""Classify each bank transaction for an Indian household expense tracker.

Types: {type_values}
Headings (for expense/investment only): {heading_values}

Rules:
- 'transfer' = NEFT/IMPS/UPI between own accounts
- 'investment' = loan EMI, SIP, mutual fund, insurance premium
- 'income' = salary credit, interest, dividend, refund
- 'official' = employer reimbursement (GCPL/Godrej)
- 'expense' = all other debits

Transactions: {json.dumps(items, indent=2)}

Reply with JSON array, same order: [{{"id":"...", "type":"...", "heading":"..." or null, "paid_to":"<clean vendor name>", "certain":true/false}}]
Only output JSON, no explanation."""

        try:
            resp = client.chat.completions.create(
                model=deployment, max_tokens=1500,
                messages=[
                    {"role":"system","content":"Reply only with JSON."},
                    {"role":"user","content":prompt},
                ]
            )
            text = resp.choices[0].message.content.strip()
            if text.startswith("```"):
                text = text.split("```")[1]
                if text.startswith("json"):
                    text = text[4:]
            results = json.loads(text)

            # Apply AI results back to transactions
            result_map = {r["id"]: r for r in results}
            for txn in batch:
                r = result_map.get(txn["txn_id"])
                if not r:
                    continue
                txn["type"]    = r.get("type", txn["type"])
                txn["heading"] = r.get("heading", txn["heading"])
                if r.get("paid_to"):
                    txn["paid_to"] = r["paid_to"]
                txn["uncertain"] = not r.get("certain", False)
                txn["uncertain_fields"] = [] if r.get("certain") else txn.get("uncertain_fields", [])
                txn["confidence"] = "ai"
        except Exception as e:
            print(f"AI batch classify error: {e}")

    return transactions


# ── Gmail sync ────────────────────────────────────────────────────────────────

def _get_gmail_service():
    from googleapiclient.discovery import build
    from src.gmail_utils import get_credentials
    return build("gmail", "v1", credentials=get_credentials())


def _decode_body(msg: dict) -> str:
    """Extract readable text from email — prefers plain text, falls back to HTML stripped."""
    import base64, re as _re

    def _decode(data: str) -> str:
        return base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")

    def _parts_text(parts, mime):
        for p in parts:
            if p.get("mimeType") == mime:
                d = p.get("body", {}).get("data", "")
                if d:
                    return _decode(d)
            sub = p.get("parts", [])
            if sub:
                r = _parts_text(sub, mime)
                if r:
                    return r
        return ""

    payload = msg.get("payload", {})
    parts   = payload.get("parts", [])

    # Try plain text first
    text = _parts_text(parts, "text/plain")
    if not text:
        # Fallback: body directly
        d = payload.get("body", {}).get("data", "")
        if d:
            text = _decode(d)

    if not text:
        # Try HTML and strip tags
        html = _parts_text(parts, "text/html")
        if html:
            text = _re.sub(r"<[^>]+>", " ", html)
            text = _re.sub(r"\s+", " ", text).strip()

    return text


def _parse_icici_savings(body: str, gmail_id: str) -> Optional[dict]:
    """ICICI savings account debit/credit alert.

    Actual formats observed:
      'A purchase of Rs. 5,740.00 has been made using your Debit Card linked to
       ICICI Bank Account XX331 on 13-Jun-26. Info: VPS*COPPER CHIM.'
      'INR X has been debited from your ICICI Bank Account XX1234 on DD-MM-YY'
      'Rs. X has been credited to your ICICI Bank Account XX1234 on DD-MM-YY'
    """
    # Amount — several formats
    m_amt = (
        re.search(r"[Rr]s\.?\s*([\d,]+(?:\.\d{2})?)\s+has been (?:made|debited|paid)", body, re.I) or
        re.search(r"purchase of\s+[Rr]s\.?\s*([\d,]+(?:\.\d{2})?)", body, re.I) or
        re.search(r"INR\s*([\d,]+(?:\.\d{2})?)\s+(?:has been )?debited", body, re.I) or
        re.search(r"payment of\s+[Rr]s\s*([\d,]+(?:\.\d{2})?)", body, re.I) or
        re.search(r"[Rr]s\s*([\d,]+(?:\.\d{2})?)\s+(?:has been )?(?:debited|paid|deducted)", body, re.I)
    )
    m_cr = (
        re.search(r"INR\s*([\d,]+(?:\.\d{2})?)\s+(?:has been )?credited", body, re.I) or
        re.search(r"[Rr]s\.?\s*([\d,]+(?:\.\d{2})?)\s+has been credited", body, re.I) or
        re.search(r"credited.*?[Rr]s\.?\s*([\d,]+(?:\.\d{2})?)", body, re.I)
    )

    if not (m_amt or m_cr):
        return None

    is_debit = bool(m_amt)
    amount   = float((m_amt or m_cr).group(1).replace(",", ""))

    m_acct  = re.search(r"(?:Account|Savings Account)\s+(?:No\.?\s+)?X*(\d{3,4})", body, re.I)
    m_date  = re.search(r"on\s+(\w{3,4}\s+\d{1,2},?\s+\d{4}|\d{1,2}[-/]\w{3,9}[-/]\d{2,4}|\d{2}[-/]\d{2}[-/]\d{2,4})", body, re.I)
    m_info  = (
        re.search(r"Info[:\s]+([^\n\.]{3,60})", body, re.I) or
        re.search(r"towards\s+([A-Z][A-Za-z\s]{3,40}?)(?:\s+from|\s+on|\.|$)", body, re.I) or
        re.search(r"payment\s+(?:to|of)\s+([A-Z][A-Za-z\s]{3,40}?)(?:\s+from|\s+on|\.|$)", body, re.I)
    )

    # Normalise 3-digit suffixes to their known 4-digit account numbers
    _ACCT_ALIAS = {"331": "1331", "281": "7281"}
    raw_suffix  = m_acct.group(1) if m_acct else None
    suffix      = _ACCT_ALIAS.get(raw_suffix, raw_suffix)
    acct_no     = f"ICICI-{suffix}" if suffix else "ICICI-????"
    date_str = m_date.group(1) if m_date else datetime.now().strftime("%d/%m/%Y")
    desc     = m_info.group(1).strip() if m_info else body[:80]

    return {
        "gmail_id":        gmail_id,
        "bank":            "ICICI",
        "account":         acct_no,
        "account_type":    "savings",
        "raw_description": desc,
        "debit":           amount if is_debit else 0,
        "credit":          0 if is_debit else amount,
        "date":            date_str,
        "source":          "gmail_alert",
    }


def _parse_icici_credit_card(body: str, gmail_id: str) -> Optional[dict]:
    """ICICI credit card spend alert."""
    m_amt  = re.search(r"(?:Rs\.?|INR)\s*([\d,]+(?:\.\d{2})?)\s+(?:has been )?(?:spent|used|debited)", body, re.I)
    m_card = re.search(r"Credit Card\s+(?:ending\s+)?(?:XX|x+)?(\d{3,4})", body, re.I)
    m_date = re.search(r"on\s+(\d{2}[-/]\d{2}[-/]\d{2,4})", body, re.I)
    m_at   = re.search(r"(?:at|for)\s+([A-Za-z0-9&\s\-\.]{3,50}?)(?:\.|on\s+\d|\n)", body, re.I)

    if not m_amt:
        return None

    amount  = float(m_amt.group(1).replace(",",""))
    card_no = f"ICICI-CC-{m_card.group(1)}" if m_card else "ICICI-CC-????"
    date_str= m_date.group(1) if m_date else datetime.now().strftime("%d/%m/%Y")
    desc    = m_at.group(1).strip() if m_at else body[:80]

    return {
        "gmail_id":        gmail_id,
        "bank":            "ICICI",
        "account":         card_no,
        "account_type":    "credit_card",
        "raw_description": desc,
        "debit":           amount,
        "credit":          0,
        "date":            date_str,
        "source":          "gmail_alert",
    }


def _parse_sbi_savings(body: str, gmail_id: str) -> Optional[dict]:
    """SBI savings debit/credit alert — emails are HTML-only."""
    # Strip style blocks and tags first
    clean = re.sub(r"<style[^>]*>.*?</style>", " ", body, flags=re.S | re.I)
    clean = re.sub(r"<[^>]+>", " ", clean)
    clean = re.sub(r"\s+", " ", clean).strip()

    m_debit  = re.search(r"[Rr]s\.?\s*([\d,]+(?:\.\d{2})?)\s+(?:has been )?(?:debited|deducted)", clean, re.I)
    m_credit = re.search(r"[Rr]s\.?\s*([\d,]+(?:\.\d{2})?)\s+(?:has been )?credited", clean, re.I)
    m_acct   = re.search(r"(?:account|A/c)[^\d]*(?:XX|x+)?(\d{3,5})", clean, re.I)
    m_date   = re.search(r"(?:on|dated)\s+(\d{1,2}[-/]\w{3,9}[-/]\d{2,4}|\d{2}[-/]\d{2}[-/]\d{2,4})", clean, re.I)
    m_info   = re.search(r"(?:Info|towards|narration|remarks)[:\s]+([A-Za-z0-9*\s\-&\.\/]{3,60}?)(?:\.|  |\n|Rs)", clean, re.I)

    if not (m_debit or m_credit):
        return None

    is_debit = bool(m_debit)
    amount   = float((m_debit or m_credit).group(1).replace(",", ""))
    acct_no  = f"SBI-{m_acct.group(1)}" if m_acct else "SBI-????"
    date_str = m_date.group(1) if m_date else datetime.now().strftime("%d/%m/%Y")
    desc     = m_info.group(1).strip() if m_info else clean[:80]

    return {
        "gmail_id":        gmail_id,
        "bank":            "SBI",
        "account":         acct_no,
        "account_type":    "savings",
        "raw_description": desc,
        "debit":           amount if is_debit else 0,
        "credit":          0 if is_debit else amount,
        "date":            date_str,
        "source":          "gmail_alert",
    }


def _parse_sbi_credit_card(body: str, gmail_id: str) -> Optional[dict]:
    """SBI credit card spend alert."""
    m_amt  = re.search(r"Rs\.?\s*([\d,]+(?:\.\d{2})?)\s+(?:spent|used|debited)", body, re.I)
    m_card = re.search(r"(?:Card|SBI Card)\s+(?:ending\s+)?(?:XX|x+)?(\d{3,4})", body, re.I)
    m_date = re.search(r"(?:on|dated)\s+(\d{2}[-/]\d{2}[-/]\d{2,4})", body, re.I)
    m_at   = re.search(r"(?:at|for)\s+([A-Za-z0-9&\s\-\.]{3,50}?)(?:\.|on\s+\d|\n)", body, re.I)

    if not m_amt:
        return None

    amount  = float(m_amt.group(1).replace(",",""))
    card_no = f"SBI-CC-{m_card.group(1)}" if m_card else "SBI-CC-????"
    date_str= m_date.group(1) if m_date else datetime.now().strftime("%d/%m/%Y")
    desc    = m_at.group(1).strip() if m_at else body[:80]

    return {
        "gmail_id":        gmail_id,
        "bank":            "SBI",
        "account":         card_no,
        "account_type":    "credit_card",
        "raw_description": desc,
        "debit":           amount,
        "credit":          0,
        "date":            date_str,
        "source":          "gmail_alert",
    }


def _strip_html(html: str) -> str:
    """Strip HTML tags and normalize whitespace."""
    clean = re.sub(r"<style[^>]*>.*?</style>", " ", html, flags=re.S | re.I)
    clean = re.sub(r"<script[^>]*>.*?</script>", " ", clean, flags=re.S | re.I)
    clean = re.sub(r"<[^>]+>", " ", clean)
    clean = re.sub(r"&nbsp;", " ", clean, flags=re.I)
    clean = re.sub(r"&amp;", "&", clean, flags=re.I)
    return re.sub(r"\s+", " ", clean).strip()


def _detect_and_parse(raw_body: str, gmail_id: str) -> Optional[dict]:
    """Detect bank/account type from email body and parse.

    Uses the clean text version for pattern matching so HTML emails work too.
    """
    # Use plain text as-is; if it looks like HTML, strip tags first
    body = raw_body
    if "<html" in raw_body.lower() or "<style" in raw_body.lower():
        body = _strip_html(raw_body)

    b = body.lower()
    is_icici = "icici bank" in b or "icicibank" in b
    is_sbi   = ("state bank" in b or "sbi card" in b) and "icici" not in b
    is_kotak = "kotak" in b
    is_cc    = "credit card" in b or "sbi card" in b

    if is_icici and is_cc:
        return _parse_icici_credit_card(body, gmail_id)
    if is_icici:
        return _parse_icici_savings(body, gmail_id)
    if is_sbi and is_cc:
        return _parse_sbi_credit_card(body, gmail_id)
    if is_sbi:
        return _parse_sbi_savings(body, gmail_id)
    if is_kotak:
        return _parse_kotak(body, gmail_id)
    return None


def _parse_kotak(body: str, gmail_id: str) -> Optional[dict]:
    """Kotak Mahindra Bank alert."""
    m_dr  = re.search(r"[Rr]s\.?\s*([\d,]+(?:\.\d{2})?)\s+(?:debited|paid|Payment Gateway)", body, re.I)
    m_cr  = re.search(r"[Rr]s\.?\s*([\d,]+(?:\.\d{2})?)\s+(?:credited|received)", body, re.I)
    m_acct= re.search(r"(?:Account|A/c)[^\d]*(\d{3,5})", body, re.I)
    m_date= re.search(r"on\s+(\d{1,2}[-/]\w{3,9}[-/]\d{2,4}|\d{2}[-/]\d{2}[-/]\d{2,4})", body, re.I)
    m_info= re.search(r"(?:Info|towards|at|Merchant)[:\s]+([A-Za-z0-9*\s\-&\.]{3,50}?)(?:\.|  |\n|Rs)", body, re.I)

    if not (m_dr or m_cr):
        return None

    is_debit = bool(m_dr)
    amount   = float((m_dr or m_cr).group(1).replace(",", ""))
    acct_no  = f"KOTAK-{m_acct.group(1)}" if m_acct else "KOTAK-????"
    date_str = m_date.group(1) if m_date else datetime.now().strftime("%d/%m/%Y")
    desc     = m_info.group(1).strip() if m_info else body[:80]

    return {
        "gmail_id":        gmail_id,
        "bank":            "KOTAK",
        "account":         acct_no,
        "account_type":    "savings",
        "raw_description": desc,
        "debit":           amount if is_debit else 0,
        "credit":          0 if is_debit else amount,
        "date":            date_str,
        "source":          "gmail_alert",
    }


PROCESSED_IDS_PATH = os.path.join(DATA_DIR, "processed_gmail_ids.json")


def _get_processed_ids() -> set:
    return set(_db.load("processed_gmail_ids", default=[]))


def _mark_processed(msg_id: str):
    ids = _get_processed_ids()
    ids.add(msg_id)
    _db.save("processed_gmail_ids", list(ids))


def sync_from_gmail(days_back: int = 90, force: bool = False) -> dict:
    """
    Pull bank alert emails from Gmail (last N days).
    Deduplicates by txn_id, classifies, saves to master_ledger.json.
    Returns summary: {new, skipped, uncertain}.
    """
    try:
        service = _get_gmail_service()
    except Exception as e:
        return {"error": str(e), "new": 0, "skipped": 0, "uncertain": 0}

    if force:
        _db.save("processed_gmail_ids", [])
        _db.save("processed_statement_ids", [])
    processed_ids = _get_processed_ids()
    existing      = _load_json(LEDGER_PATH)
    existing_ids  = {t["txn_id"] for t in existing}

    # Search bank alert emails (last N days)
    # Actual senders confirmed from Gmail inspection:
    #   ICICI: alert@icici.bank.in, customernotification@icici.bank.in
    #   SBI:   yonobysbi@alerts.sbi.bank.in, sbialerts@sbi.co.in
    #   Kotak: bankalerts@kotak.bank.in (if applicable)
    cutoff = (datetime.now() - timedelta(days=days_back)).strftime("%Y/%m/%d")
    query  = (
        f"after:{cutoff} "
        "(from:alert@icici.bank.in OR from:customernotification@icici.bank.in "
        "OR from:yonobysbi@alerts.sbi.bank.in OR from:sbialerts@sbi.co.in "
        "OR from:bankalerts@kotak.bank.in OR from:alerts@sbi.co.in "
        "OR from:sbicardservices@sbi.co.in OR from:info@sbicard.com) "
        "(debited OR credited OR purchase OR \"transaction alert\")"
    )

    new_txns = []
    skipped  = 0

    try:
        result = service.users().messages().list(
            userId="me", q=query, maxResults=500
        ).execute()
        messages = result.get("messages", [])
    except Exception as e:
        return {"error": str(e), "new": 0, "skipped": 0, "uncertain": 0}

    for msg_ref in messages:
        msg_id = msg_ref["id"]
        if msg_id in processed_ids:
            skipped += 1
            continue

        try:
            msg  = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
            body = _decode_body(msg)
            parsed = _detect_and_parse(body, msg_id)

            if not parsed:
                _mark_processed(msg_id)
                continue

            # Build full transaction record
            dt = _parse_date(parsed["date"])
            if not dt:
                dt = datetime.now()
            date_str = dt.strftime("%d/%m/%Y")

            fy = _fy_info(dt)
            txn_id = _txn_id(
                date_str, parsed["account"],
                parsed["raw_description"], parsed["debit"] or parsed["credit"]
            )

            if txn_id in existing_ids:
                skipped += 1
                _mark_processed(msg_id)
                continue

            txn = {
                "txn_id":          txn_id,
                "date":            date_str,
                "fy_month_no":     fy["fy_month_no"],
                "fy_month_name":   fy["fy_month_name"],
                "fy_year":         fy["fy_year"],
                "account":         parsed["account"],
                "account_type":    parsed["account_type"],
                "bank":            parsed["bank"],
                "raw_description": parsed["raw_description"],
                "paid_to":         "",
                "debit":           parsed["debit"],
                "credit":          parsed["credit"],
                "type":            "",
                "heading":         None,
                "remarks":         "",
                "uncertain":       True,
                "uncertain_fields": [],
                "confidence":      "new",
                "source":          "gmail_alert",
                "gmail_id":        parsed["gmail_id"],
                "ai_saving_tip":   None,
                "saving_agreed":   None,
                "reconciled_with": None,
                "created_at":      datetime.now().isoformat(),
            }

            txn = classify_transaction(txn)
            new_txns.append(txn)
            existing_ids.add(txn_id)
            _mark_processed(msg_id)

        except Exception as e:
            print(f"Error processing Gmail msg {msg_id}: {e}")

    if new_txns:
        # AI classify uncertain transactions
        new_txns = ai_classify_batch(new_txns)
        existing.extend(new_txns)
        # Sort by date descending
        existing.sort(key=lambda t: t.get("date",""), reverse=True)
        _save_json(LEDGER_PATH, existing)

    uncertain_count = sum(1 for t in existing if t.get("uncertain"))
    return {"new": len(new_txns), "skipped": skipped, "uncertain": uncertain_count}


# ── Reconcile with approval log ───────────────────────────────────────────────

def reconcile_with_approvals(approval_log: list) -> int:
    """
    Match SBI debits in master ledger to Vincent's approved expenses.
    Uses heading from approval log when matched. Returns count of new matches.
    """
    ledger = _load_json(LEDGER_PATH)
    matched = 0

    for txn in ledger:
        if txn.get("reconciled_with") or txn.get("bank") != "SBI":
            continue
        if txn.get("account_type") == "credit_card":
            continue

        txn_date = _parse_date(txn["date"])
        txn_amt  = txn["debit"]
        if not txn_date or not txn_amt:
            continue

        for entry in approval_log:
            if entry.get("action") not in ("AUTO_APPROVE","APPROVED","APPROVED_LOWER"):
                continue
            if entry.get("reconciled_to"):
                continue

            log_amt = entry.get("approved_amount") or entry.get("amount", 0)
            if abs(txn_amt - log_amt) > 100:
                continue

            try:
                log_date = datetime.fromisoformat(entry["timestamp"])
            except Exception:
                continue

            if abs((txn_date - log_date).days) > 7:
                continue

            # Match found — pull heading from approval log
            from src.approval_engine import ApprovalEngine
            APP_TO_HEADING = {
                "groceries":"Groceries","staff":"Staff Salary","utilities":"Electricity & Gas",
                "miscellaneous":"Misc","personal_care":"Wellness","clothing":"Clothes",
                "gifts":"Gifts","medical":"Medical","education":"Children Education",
                "dining":"Eating Out","entertainment":"Entertainment","transport":"Holiday",
                "maintenance":"Maintenance Expense","home_repair":"One Time Charge",
            }
            cat = entry.get("category","miscellaneous")
            heading = APP_TO_HEADING.get(cat, txn.get("heading","Misc"))

            txn["heading"]         = heading
            txn["reconciled_with"] = entry["request_id"]
            txn["uncertain"]       = False
            txn["uncertain_fields"]= []
            txn["confidence"]      = "reconciled"
            entry["reconciled_to"] = txn["txn_id"]
            matched += 1
            break

    if matched:
        _save_json(LEDGER_PATH, ledger)

    return matched


# ── Credit card balance tracking ──────────────────────────────────────────────

def get_cc_balance() -> dict:
    """Return unpaid credit card balance per card."""
    ledger = _load_json(LEDGER_PATH)
    balances: dict = {}

    for txn in ledger:
        if txn.get("account_type") != "credit_card":
            continue
        card = txn["account"]
        if card not in balances:
            balances[card] = {"card": card, "unpaid": 0, "total_spend": 0, "txns": []}
        balances[card]["total_spend"] += txn.get("debit", 0)

        # Credit = payment made, debit = spend
        if txn.get("credit"):
            balances[card]["unpaid"] -= txn["credit"]
        else:
            balances[card]["unpaid"] += txn.get("debit", 0)

        balances[card]["txns"].append(txn["txn_id"])

    return balances


# ── Public API functions ──────────────────────────────────────────────────────

def deduplicate_ledger() -> int:
    """Remove duplicate transactions caused by same txn appearing in Gmail alert + PDF.
    Groups by (date, account, debit, credit). Keeps the richest entry, merges fields.
    Returns number of duplicates removed.
    """
    from collections import defaultdict
    ledger = _load_json(LEDGER_PATH)

    def _score(t):
        s = 0
        if t.get("confidence") == "manual":  s += 100
        if t.get("raw_description"):          s += 20
        if t.get("paid_to"):                  s += 10
        if t.get("type") and t["type"] not in ("", "Expense", "expense"): s += 5
        if t.get("heading") and t["heading"] not in ("Unknown", "Misc", "", None): s += 5
        if t.get("source") == "pdf_import":   s += 2
        return s

    groups = defaultdict(list)
    for txn in ledger:
        key = (
            txn.get("date", ""),
            txn.get("account", ""),
            round(float(txn.get("debit") or 0), 2),
            round(float(txn.get("credit") or 0), 2),
        )
        groups[key].append(txn)

    to_remove = set()
    for key, txns in groups.items():
        if len(txns) <= 1:
            continue
        txns.sort(key=_score, reverse=True)
        winner = txns[0]
        for loser in txns[1:]:
            # Merge useful fields the winner lacks
            if not winner.get("raw_description") and loser.get("raw_description"):
                winner["raw_description"] = loser["raw_description"]
            if not winner.get("paid_to") and loser.get("paid_to"):
                winner["paid_to"] = loser["paid_to"]
            if not winner.get("gmail_id") and loser.get("gmail_id"):
                winner["gmail_id"] = loser["gmail_id"]
            if (not winner.get("type") or winner["type"] in ("", "Unknown")) and loser.get("type"):
                winner["type"] = loser["type"]
            if (not winner.get("heading") or winner["heading"] in ("", "Unknown", "Misc")) and loser.get("heading"):
                winner["heading"] = loser["heading"]
            to_remove.add(loser["txn_id"])

    if to_remove:
        ledger = [t for t in ledger if t["txn_id"] not in to_remove]
        _save_json(LEDGER_PATH, ledger)
    return len(to_remove)


def import_from_icici_transactions() -> int:
    """
    One-time import: pull records from icici_transactions.json into master ledger.
    Deduplicates by txn_id and by (date, account, amount) to prevent cross-source dupes.
    Returns count of newly added records.
    """
    existing = _load_json(LEDGER_PATH)
    existing_ids = {t["txn_id"] for t in existing}
    # Secondary dedup key: same transaction already in ledger from another source
    existing_dupe_keys = {
        (t.get("date",""), t.get("account",""),
         round(float(t.get("debit") or 0), 2),
         round(float(t.get("credit") or 0), 2))
        for t in existing
    }

    source = _db.load("icici_transactions")
    added = 0

    for t in source:
        txn_id = t.get("txn_id")
        if not txn_id or txn_id in existing_ids:
            continue

        raw_date = t.get("date", "")
        dt = _parse_date(raw_date)
        if not dt:
            dt = datetime.now()
        dt_str = dt.strftime("%d/%m/%Y")
        fy = _fy_info(dt)

        debit  = t.get("debit") or (t.get("amount", 0) if t.get("type") == "debit" else 0)
        credit = t.get("credit") or (t.get("amount", 0) if t.get("type") == "credit" else 0)

        dupe_key = (dt_str, t.get("account", ""), round(float(debit), 2), round(float(credit), 2))
        if dupe_key in existing_dupe_keys:
            continue

        txn = {
            "txn_id":          txn_id,
            "date":            dt_str,
            "fy_month_no":     fy["fy_month_no"],
            "fy_month_name":   fy["fy_month_name"],
            "fy_year":         fy["fy_year"],
            "account":         t.get("account", "ICICI-????"),
            "account_type":    "savings",
            "bank":            "ICICI",
            "raw_description": t.get("transaction_details") or t.get("description", ""),
            "paid_to":         t.get("paid_to", ""),
            "debit":           debit,
            "credit":          credit,
            "type":            t.get("acc_type", ""),
            "heading":         t.get("heading"),
            "remarks":         t.get("remarks", ""),
            "uncertain":       t.get("confidence", "rule") != "manual",
            "uncertain_fields": [],
            "confidence":      t.get("confidence", "rule"),
            "source":          "pdf_import",
            "gmail_id":        None,
            "ai_saving_tip":   None,
            "saving_agreed":   None,
            "reconciled_with": None,
            "created_at":      datetime.now().isoformat(),
        }

        # Re-classify if type/heading not set
        if not txn["type"]:
            txn = classify_transaction(txn)

        existing.append(txn)
        existing_ids.add(txn_id)
        existing_dupe_keys.add(dupe_key)
        added += 1

    if added:
        existing.sort(key=lambda t: t.get("date", ""), reverse=True)
        _save_json(LEDGER_PATH, existing)

    return added


def repair_pdf_descriptions() -> int:
    """
    Backfill raw_description and classification from icici_transactions into
    ledger entries that have blank descriptions (imported before the field-name fix).
    Returns count of repaired entries.
    """
    ledger = _load_json(LEDGER_PATH)

    # Normalise legacy / truncated account names
    _ACCT_RE    = re.compile(r"^icic(\d{3,4})$", re.IGNORECASE)
    _ALIAS_MAP  = {"331": "1331", "281": "7281"}   # 3-digit → 4-digit canonical

    def _fix_account(name: str) -> str:
        m = _ACCT_RE.match(name or "")
        if m:
            sfx = m.group(1)
            return f"ICICI-{_ALIAS_MAP.get(sfx, sfx)}"
        sfx = _ALIAS_MAP.get((name or "").replace("ICICI-", ""))
        if sfx:
            return f"ICICI-{sfx}"
        return name

    acct_fixed = 0
    for txn in ledger:
        fixed = _fix_account(txn.get("account", ""))
        if fixed != txn.get("account"):
            txn["account"] = fixed
            acct_fixed += 1
    if acct_fixed:
        _save_json(LEDGER_PATH, ledger)

    icici_txns = _db.load("icici_transactions")
    icici_fixed = 0
    for t in icici_txns:
        fixed = _fix_account(t.get("account", ""))
        if fixed != t.get("account"):
            t["account"] = fixed
            icici_fixed += 1
    if icici_fixed:
        _db.save("icici_transactions", icici_txns)

    source_map = {t.get("txn_id"): t for t in icici_txns if t.get("txn_id")}

    # Fallback index by (date, amount, account) for when txn_id hash changed
    fallback_map = {}
    for t in icici_txns:
        key = (t.get("date", ""), str(t.get("amount", "")), t.get("account", ""))
        fallback_map.setdefault(key, t)

    repaired = 0
    for txn in ledger:
        if txn.get("source") != "pdf_import" or txn.get("confidence") == "manual":
            continue
        needs_desc = not txn.get("raw_description")
        needs_class = not txn.get("type") or not txn.get("heading") or txn.get("heading") in ("Unknown", "")
        if not needs_desc and not needs_class:
            continue

        # Try txn_id match first, fall back to (date, amount, account)
        src = source_map.get(txn["txn_id"])
        if not src:
            amount = txn.get("debit") or txn.get("credit") or 0
            fkey = (txn.get("date", ""), str(float(amount)), txn.get("account", ""))
            src = fallback_map.get(fkey)

        if needs_desc:
            if src:
                txn["raw_description"] = src.get("transaction_details") or src.get("description", "")
            # 7281 OD account: if still no description but we can infer from amount context, leave for rule below

        if not txn.get("paid_to") and src and src.get("paid_to"):
            txn["paid_to"] = src["paid_to"]

        if needs_class:
            desc = txn.get("raw_description", "")
            if desc:
                classified = classify_transaction(dict(txn))
                txn["type"]    = classified.get("type") or txn.get("type", "")
                txn["heading"] = classified.get("heading") or txn.get("heading")
                txn["paid_to"] = txn.get("paid_to") or classified.get("paid_to", "")
            elif txn.get("account") == "ICICI-7281":
                # 7281 OD account with no description: default to Financial Expense
                # (all 7281 charges are either OD interest or bank fees)
                txn["type"]    = "Expense"
                txn["heading"] = "Financial Expense"
                txn["uncertain"] = True

        repaired += 1

    # Second pass: reclassify ANY non-manual entry with a description but bad heading
    # (catches gmail_alert entries that were classified before rule updates)
    BAD_HEADINGS = {"Unknown", "Misc", "", None}
    reclassified = 0
    for txn in ledger:
        if txn.get("confidence") == "manual":
            continue
        if txn.get("heading") not in BAD_HEADINGS:
            continue
        desc = txn.get("raw_description", "")
        if not desc:
            continue
        classified = classify_transaction(dict(txn))
        new_type    = classified.get("type", "")
        new_heading = classified.get("heading", "")
        if new_type and new_heading not in BAD_HEADINGS:
            txn["type"]    = new_type
            txn["heading"] = new_heading or ""
            txn["paid_to"] = txn.get("paid_to") or classified.get("paid_to", "")
            txn["uncertain"] = not classified.get("uncertain", True)
            reclassified += 1

    # Third pass: round-thousand unclassified amounts + sudhir sitapati payee = Transfer
    # Only applies when heading is still bad (Unknown/Misc/empty) to avoid overriding
    # known categories like Staff Salary, Home Loan, Charity which are also round numbers.
    round_fixed = 0
    for txn in ledger:
        if txn.get("confidence") == "manual":
            continue
        paid_to = (txn.get("paid_to") or "").lower()
        desc    = (txn.get("raw_description") or "").lower()
        is_self = "sudhir sitapati" in paid_to or "sudhir sitapati" in desc or "sudhirsitapati" in desc
        debit   = float(txn.get("debit") or 0)
        credit  = float(txn.get("credit") or 0)
        amount  = debit or credit
        is_round_k = amount >= 1000 and amount % 1000 == 0
        heading = txn.get("heading") or ""
        is_bad  = heading in BAD_HEADINGS
        if (is_self or (is_round_k and is_bad)) and txn.get("type") not in ("Transfer", "Income", "Investment", "Official"):
            txn["type"]    = "Transfer"
            txn["heading"] = "Interbank"
            txn["uncertain"] = False
            round_fixed += 1

    if repaired or reclassified or round_fixed:
        _save_json(LEDGER_PATH, ledger)
    return repaired + reclassified + round_fixed


def load_ledger() -> list:
    return _load_json(LEDGER_PATH)


def get_uncertain() -> list:
    return [t for t in load_ledger() if t.get("uncertain")]


def update_transaction(txn_id: str, updates: dict) -> bool:
    """Update editable fields. Clears uncertain flag if type+heading provided."""
    ledger = _load_json(LEDGER_PATH)
    allowed = {"paid_to","type","heading","remarks","ai_saving_tip","saving_agreed"}
    for txn in ledger:
        if txn["txn_id"] == txn_id:
            for k, v in updates.items():
                if k in allowed:
                    txn[k] = v
            if "type" in updates or "heading" in updates:
                txn["confidence"] = "manual"
                txn["uncertain"]  = False
                txn["uncertain_fields"] = []
            _save_json(LEDGER_PATH, ledger)
            return True
    return False
