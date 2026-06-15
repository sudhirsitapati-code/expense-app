"""
icici_classifier.py
Rule-based + AI classification of ICICI/SBI transactions into ACC26 fields.
Rules derived from SudhirExpenses ACC26ver5_MASTER.xlsx.

Type values:  Expense | Investment | Transfer | Income | Official | Error
Heading values match ExpenseSummary exactly (see ACC26 for canonical list).
"""

import json
import os
from openai import AzureOpenAI


# ── ACC26 canonical headings (from ExpenseSummary) ────────────────────────────

HEADINGS = [
    # HOUSEHOLD
    "Groceries", "Staff Salary", "Electricity & Gas", "Misc", "Cash",
    # PERSONAL
    "Alcohol", "Wellness",
    # FAMILY
    "Clothes", "Gifts", "Medical", "Amma", "Ketki", "Children Education",
    # GIVING
    "Charity", "Uspaar",
    # LIFESTYLE
    "Holiday", "Eating Out", "Entertainment",
    # PROPERTY
    "Malhar", "Maintenance Expense", "Home office", "One Time Charge",
    "Kalpataru Maintenance",
    # FINANCIAL
    "Financial Expense", "Insurance", "Home Loan", "Tax",
    # INVESTMENTS / TRANSFERS
    "Art", "Foreign Investment", "GCPL Share Sale", "Property Investment",
    "Loan Repayment",
    # OTHER
    "Interbank", "Salary", "Refund", "Unknown",
]


# ── Keyword rules — order matters (first match wins) ─────────────────────────
# Each rule: list of substrings (lowercase, any match) → (Type, Heading)

RULES = [
    # ── INCOME ────────────────────────────────────────────────────────────────
    (["gcplsalary", "godrej consumer products ltd sal", "godrej consumer salary",
      "cms/gcpl", "salary credit", "sal/cr"],
     "Income", "Salary"),

    (["interim dividend", "gcpl interim", "dividend"],
     "Income", "Salary"),           # treat dividends under Income; heading flexible

    # ── OFFICIAL ──────────────────────────────────────────────────────────────
    (["neft-hsbcn", "gcpl reimburs", "official reimburs", "boss petrol bill",
      "transfer cridit card boss"],
     "Official", "Financial Expense"),

    # ── TAX ───────────────────────────────────────────────────────────────────
    (["dtax", "gib/00", "income tax", "advance tax", "tcs on lrs",
      "trf/godrej consumer products ltd esgs", "taxs"],
     "Expense", "Tax"),

    # ── HOME LOAN / OD INTEREST ───────────────────────────────────────────────
    (["bil/home loan", "home loan xx", "emi sudhir", "xx99508", "xx00382",
      "smp/tbmum", "102205009175"],
     "Expense", "Home Loan"),

    # ── INSURANCE ─────────────────────────────────────────────────────────────
    (["life insurance corporation", "trf/lic", "insurance", "policy"],
     "Expense", "Insurance"),

    # ── INVESTMENTS ───────────────────────────────────────────────────────────
    (["mizugami", "art gallery"],
     "Investment", "Art"),

    (["nrs/usd", "foreign invest", "lrs remittance"],
     "Investment", "Foreign Investment"),

    (["eba/eq trade", "gcpl share", "gcpl eq", "eq trade"],
     "Investment", "GCPL Share Sale"),

    (["kamala ganesh", "trfr to kamala", "kashid"],
     "Investment", "Property Investment"),

    (["infina finance", "neft.*infina", "loan repay", "kotak mahindra prime"],
     "Investment", "Loan Repayment"),

    (["uspaar"],
     "Investment", "Uspaar"),

    # ── TRANSFER (credit card / interbank) ────────────────────────────────────
    (["icici bank credit ca", "bil/001", "credit card", "552418",
      "credit ca/", "interbank"],
     "Transfer", "Interbank"),

    (["trfr to", "transfer to icici", "transfer to hdfc", "transfer to sbi",
      "neft/", "rtgs/", "imps/"],
     "Transfer", "Interbank"),

    # ── KALPATARU MAINTENANCE ─────────────────────────────────────────────────
    (["bmc/ici", "kalpataru maint", "society maintenance", "hsg soc",
      "kalpataru soc"],
     "Expense", "Kalpataru Maintenance"),

    # ── CHARITY ───────────────────────────────────────────────────────────────
    (["kalpalata", "dhamma patt", "durga kavitha", "priyanka londhe",
      "donation", "charity"],
     "Expense", "Charity"),

    # ── STAFF SALARY ─────────────────────────────────────────────────────────
    (["vincent fe", "vincent salary", "santosh", "mary", "shiloj", "mohammad",
      "staff salary", "maid salary", "cook salary", "driver salary",
      "transfer mary", "transfer vincent", "neft.*vincent"],
     "Expense", "Staff Salary"),

    # ── CHILDREN EDUCATION ────────────────────────────────────────────────────
    (["oberoi school", "oberoi internat", "tridha school", "school fees",
      "sahaana", "kabir", "tuition", "dhirubhai ambani", "das"],
     "Expense", "Children Education"),

    # ── AMMA ──────────────────────────────────────────────────────────────────
    (["amma", "kamala sitapati"],
     "Expense", "Amma"),

    # ── KETKI ─────────────────────────────────────────────────────────────────
    (["ketki", "kiran sitapati"],
     "Expense", "Ketki"),

    # ── HOLIDAY ───────────────────────────────────────────────────────────────
    (["thomas cook", "makemytrip", "irctc", "holiday", "air india",
      "indigo", "vistara", "goair", "spicejet", "hotel", "resort",
      "abi saad", "vps/abi"],
     "Expense", "Holiday"),

    # ── WELLNESS ──────────────────────────────────────────────────────────────
    (["bombay gymkhaana", "gymkhana", "gym", "yoga", "wellness",
      "fitness", "spa"],
     "Expense", "Wellness"),

    # ── EATING OUT ────────────────────────────────────────────────────────────
    (["swiggy", "zomato", "restaurant", "cafe", "pizza", "mcdonalds",
      "kfc", "starbucks", "dining"],
     "Expense", "Eating Out"),

    # ── ENTERTAINMENT ─────────────────────────────────────────────────────────
    (["netflix", "amazon prime", "spotify", "hotstar", "bookmyshow",
      "pvr", "inox", "entertainment"],
     "Expense", "Entertainment"),

    # ── ELECTRICITY & GAS ─────────────────────────────────────────────────────
    (["tata power", "msedcl", "best electricity", "mahanagar gas", "mgl",
      "electricity", "gas bill", "transfer tata power"],
     "Expense", "Electricity & Gas"),

    # ── GROCERIES ─────────────────────────────────────────────────────────────
    (["bigbasket", "blinkit", "grofers", "nature's basket", "dmart",
      "reliance fresh", "more retail", "grocery", "supermarket"],
     "Expense", "Groceries"),

    # ── MEDICAL ───────────────────────────────────────────────────────────────
    (["hospital", "clinic", "pharmacy", "medplus", "apollo", "lilavati",
      "kokilaben", "hinduja", "breach candy", "medicine", "doctor",
      "medical", "health"],
     "Expense", "Medical"),

    # ── CLOTHES ───────────────────────────────────────────────────────────────
    (["zara", "h&m", "gap", "myntra", "ajio", "lifestyle store",
      "nykaa fashion", "clothes", "clothing", "tailoring"],
     "Expense", "Clothes"),

    # ── GIFTS ─────────────────────────────────────────────────────────────────
    (["gift", "present", "amazon.*gift", "flipkart.*gift"],
     "Expense", "Gifts"),

    # ── ALCOHOL ───────────────────────────────────────────────────────────────
    (["wine", "beer", "liquor", "alcohol", "spirits", "whisky", "hilife",
      "wb's", "living liquidz"],
     "Expense", "Alcohol"),

    # ── MALHAR (farm/second property) ─────────────────────────────────────────
    (["malhar", "satoshi", "farm"],
     "Expense", "Malhar"),

    # ── HOME OFFICE ───────────────────────────────────────────────────────────
    (["home office", "office supply", "printer", "laptop", "computer",
      "stationery"],
     "Expense", "Home office"),

    # ── MAINTENANCE / HOME REPAIR ─────────────────────────────────────────────
    (["plumber", "carpenter", "electrician", "pest control", "painting",
      "repair", "maintenance", "ac service", "deep clean"],
     "Expense", "Maintenance Expense"),

    # ── MISC / CASH ATM ───────────────────────────────────────────────────────
    (["atm cash", "atm withdrawal", "cash withdrawal"],
     "Expense", "Cash"),
]


def _rule_classify(description: str):
    """Try keyword rules. Returns (type, heading) or None."""
    desc_lower = description.lower()
    for keywords, acc_type, heading in RULES:
        if any(kw in desc_lower for kw in keywords):
            return acc_type, heading
    return None


def _ai_classify_batch(transactions: list, client: AzureOpenAI, deployment: str) -> list:
    if not transactions:
        return []

    items = "\n".join(
        f"{i+1}. {t.get('description','')} | "
        f"{'Dr' if t.get('debit',0) else 'Cr'} Rs {t.get('debit') or t.get('credit',0):,.0f}"
        for i, t in enumerate(transactions)
    )

    prompt = f"""You are classifying bank statement transactions for Sudhir Sitapati, a senior executive at Godrej Consumer Products Ltd, living in Mumbai (Kalpataru building).

His expense categories (from his accounts workbook):
HOUSEHOLD: Groceries, Staff Salary, Electricity & Gas, Misc, Cash
PERSONAL: Alcohol, Wellness
FAMILY: Clothes, Gifts, Medical, Amma, Ketki, Children Education
GIVING: Charity, Uspaar
LIFESTYLE: Holiday, Eating Out, Entertainment
PROPERTY: Malhar, Maintenance Expense, Home office, One Time Charge, Kalpataru Maintenance
FINANCIAL: Financial Expense, Insurance, Home Loan, Tax
INVESTMENTS: Art, Foreign Investment, GCPL Share Sale, Property Investment, Loan Repayment
OTHER: Interbank, Salary, Refund, Unknown

Type values: Expense | Investment | Transfer | Income | Official | Error

Classify each transaction:
{items}

Reply ONLY with a JSON array of {len(transactions)} objects:
[{{"type":"Expense","heading":"Groceries","paid_to":"Reliance Fresh"}}, ...]
paid_to: short payee name if identifiable, else null."""

    try:
        response = client.chat.completions.create(
            model=deployment,
            max_tokens=600,
            messages=[
                {"role": "system", "content": "Reply only with a JSON array. No markdown."},
                {"role": "user", "content": prompt}
            ]
        )
        text = response.choices[0].message.content.strip()
        if text.startswith("```"):
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        return json.loads(text)
    except Exception as e:
        print(f"AI classify error: {e}")
        return [{"type": "Expense", "heading": "Unknown", "paid_to": None}] * len(transactions)


def classify_transactions(transactions: list) -> list:
    """
    Classify each transaction in-place, adding acc_type, heading, paid_to, confidence.
    Rule-based first (~80% coverage), AI fallback for the rest (batched in 10s).
    """
    ai_needed_idx = []
    ai_needed_txns = []

    for i, txn in enumerate(transactions):
        result = _rule_classify(txn.get("description", ""))
        if result:
            txn["acc_type"], txn["heading"] = result
            txn.setdefault("paid_to", None)
            txn["confidence"] = "high"
        else:
            ai_needed_idx.append(i)
            ai_needed_txns.append(txn)

    if not ai_needed_txns:
        return transactions

    client = AzureOpenAI(
        api_key=os.getenv("AZURE_OPENAI_KEY"),
        azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
        api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-12-01-preview"),
    )
    deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-5.5")

    BATCH = 10
    for start in range(0, len(ai_needed_txns), BATCH):
        batch = ai_needed_txns[start:start + BATCH]
        results = _ai_classify_batch(batch, client, deployment)
        for j, r in enumerate(results):
            idx = ai_needed_idx[start + j]
            transactions[idx]["acc_type"] = r.get("type", "Expense")
            transactions[idx]["heading"] = r.get("heading", "Unknown")
            transactions[idx]["paid_to"] = r.get("paid_to")
            transactions[idx]["confidence"] = "ai"

    return transactions
