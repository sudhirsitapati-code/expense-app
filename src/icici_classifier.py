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
    # ── SALARY (Income) ───────────────────────────────────────────────────────
    (["cms/ gcplsalary", "cms/gcplsalary", "godrej consumer products ltd sal",
      "gcplsalarymay", "gcplsalaryjun", "gcplsalaryjul", "gcplsalaryaug",
      "gcplsalarysep", "gcplsalaryoct", "gcplsalarynov", "gcplsalarydec",
      "gcplsalaryjan", "gcplsalaryfeb", "gcplsalarymar", "gcplsalaryapr",
      "salary credit", "sal/cr"],
     "Income", "Salary"),

    (["interim dividend", "gcpl interim", "dividend",
      "clg/vita technology", "clg/latent advisors"],
     "Income", "Salary"),

    # ── OFFICIAL (GCPL reimbursements) ────────────────────────────────────────
    (["neft-hsbcn", "gcpl reimburs", "official reimburs",
      "cms/ reimbursement pyt/godrej"],
     "Official", "Financial Expense"),

    # ── TAX ───────────────────────────────────────────────────────────────────
    (["dtax", "gib/00", "gib/002", "tcs on lrs", "tcs/", "/dtax",
      "trf/godrej consumer products ltd esgs", "itdtax",
      "advance tax", "income tax"],
     "Expense", "Tax"),

    # ── OD / BANK INTEREST & CHARGES ─────────────────────────────────────────
    (["int.coll", "int coll", "interest collected", "renewal fees",
      "sgst", "cgst", "bank charges", "service charge", "processing fee",
      "transfer int"],
     "Expense", "Financial Expense"),

    # ── HOME LOAN ─────────────────────────────────────────────────────────────
    # 003801011331:Int.Pd is interest EARNED on savings (credit), not home loan
    (["bil/home loan", "home loan xx", "emi sudhir",
      "xx99508", "xx00382", "xx42596",
      "102205009175", "tbmum0000749", "tbmum0000750", "smp/tbmum",
      "cms/ tbmum"],
     "Expense", "Home Loan"),

    # ── INTEREST INCOME (savings account credit) ──────────────────────────────
    (["011331:int.pd", "int.pd:"],
     "Income", "Salary"),

    # ── INSURANCE ─────────────────────────────────────────────────────────────
    (["life insurance corporation", "trf/lic", "trf/life insurance",
      "care health ins", "oriental insurance", "insurance", "policy"],
     "Expense", "Insurance"),

    # ── ART ───────────────────────────────────────────────────────────────────
    (["mizugami", "art gallery"],
     "Investment", "Art"),

    # ── FOREIGN INVESTMENT ────────────────────────────────────────────────────
    (["nrs/usd", "foreign invest", "lrs remittance"],
     "Investment", "Foreign Investment"),

    # ── GCPL SHARE SALE ───────────────────────────────────────────────────────
    (["eba/eq trade", "gcpl share", "gcpl eq", "eq trade"],
     "Investment", "GCPL Share Sale"),

    # ── PROPERTY INVESTMENT ───────────────────────────────────────────────────
    (["kamala ganesh", "trfr to: kamala", "trfr to:kamala", "kashid"],
     "Investment", "Property Investment"),

    # ── LOAN REPAYMENT ────────────────────────────────────────────────────────
    (["infina finance", "kotak mahindra prime", "sgb/1931",
      "nishad yogesh kapadia", "vinay ganesh sitapati"],
     "Investment", "Loan Repayment"),

    # ── USPAAR ────────────────────────────────────────────────────────────────
    (["uspaar"],
     "Investment", "Uspaar"),

    # ── SELF-TRANSFER / INTERBANK ─────────────────────────────────────────────
    # INF/INFT and NF/INFT = ICICI internet funds transfer between own accounts
    (["inf/inft", "nf/inft", "/inft/", "inft/", "net banking inf",
      "bil/001051", "bil/001105", "icici bank credit ca", "552418",
      "rtgs/icicr1", "bil/neft/in126", "bil/neft/icicn",
      "sudhir sit/state", "sudhir sit/sbin",
      "neft/", "rtgs/", "imps/", ":transfer ", "007281:", "011331:",
      "bil/neft", "bil/imps", "/inf/", "trfr to",
      "@okaxis", "@oksbi", "@okicici", "@ybl", "@ibl",
      "fund transfer"],
     "Transfer", "Interbank"),

    # ── WELLNESS ──────────────────────────────────────────────────────────────
    (["susan.9510@wah", "susan.9510@wahd", "sue.j.walker",
      "mca-shirke", "bombay gymkhaana", "gymkhana",
      "gym", "yoga", "wellness", "fitness", "spa", "pilates"],
     "Expense", "Wellness"),

    # ── STAFF SALARY ──────────────────────────────────────────────────────────
    (["vincent fe", "vincent fern", "vincent salary",
      "shiloj", "shiloch", "mohammad", "mohamad",
      "staff salary", "maid salary", "cook salary", "driver salary"],
     "Expense", "Staff Salary"),

    # ── KETKI ─────────────────────────────────────────────────────────────────
    (["ketkisitap", "ketki sitapati", "ketki"],
     "Expense", "Ketki"),

    # ── AMMA ──────────────────────────────────────────────────────────────────
    (["amma", "kamala sitapati"],
     "Expense", "Amma"),

    # ── CHILDREN EDUCATION ────────────────────────────────────────────────────
    (["oberoi school", "oberoi internat", "oberoifo", "tridha school",
      "school fees", "sahaana", "kabir", "tuition", "dhirubhai ambani",
      "komon classes", "roshini teacher", "grayquest"],
     "Expense", "Children Education"),

    # ── KALPATARU MAINTENANCE ─────────────────────────────────────────────────
    (["trf/bmc/ici", "bmc/ici", "kalpataru maint", "society maintenance",
      "hsg soc", "kalpataru soc", "kalpataru"],
     "Expense", "Kalpataru Maintenance"),

    # ── CHARITY ───────────────────────────────────────────────────────────────
    (["dhamma patt", "kalpalata", "st catherines", "donation", "charity",
      "kasaresantosh"],
     "Expense", "Charity"),

    # ── HOLIDAY ───────────────────────────────────────────────────────────────
    (["vps/abi saad", "abi saad fr", "vps/confiserie", "vps/raclette",
      "vps/hang s viet", "vps/airport tax", "vps/smiggle",
      "vps/the seraya", "vps/yatt", "gabbars bus", "bukhor",
      "vps/gamewatcher", "ips/mumbai trav",
      "wallwood ga", "vivanta", "neemrana", "alibaug",
      "makemytrip", "thomas cook", "irctc",
      "air india", "indigo", "vistara", "goair", "spicejet",
      "hotel", "resort", "airport tax",
      "kloten", "zurich", "luzern", "jbeil", "singapore",
      "manggarai", "bhubanesh", "coonoor", "coimbator",
      "tashkent", "uzbek", "yatt sadyko", "yatt tulaye",
      "vat/cash wdl", "nfs/cash wdl"],
     "Expense", "Holiday"),

    # ── EATING OUT ────────────────────────────────────────────────────────────
    (["vps/bkc saz", "vps/forty two", "vps/jap restaur", "vps/eih limited",
      "vps/shree thake", "vps/shree ukhs", "vps/smaaash mum",
      "vps/status rest", "vps/taj lands e", "vps/westin mumb",
      "vps/buono pizze", "vps/mayfairhot", "vps/msw mahesh",
      "vps/msw rigmor", "vps/gravity", "vps/nmacc food",
      "vps/semolina", "vps/madras crea", "vps/carnatic",
      "vps/manis cafe", "vps/mansuri cat",
      "vps/bombay coff", "vps/tata starbu", "vps/c254 costa",
      "swiggy", "zomato", "lin/m s profili", "vyapar.17140743",
      "restaurant", "cafe", "pizza", "mcdonalds", "kfc", "dining"],
     "Expense", "Eating Out"),

    # ── ENTERTAINMENT ─────────────────────────────────────────────────────────
    (["netflixupi", "netflix", "vps/pvr limited", "vps/kitab",
      "vps/april moon", "bookmyshow", "amazon prime", "spotify",
      "hotstar", "inox", "jioworld", "kedar.teny",
      "apple medi", "appleservices", "applamp", "entertainment"],
     "Expense", "Entertainment"),

    # ── GROCERIES ─────────────────────────────────────────────────────────────
    (["bigbasket", "blinkit", "grofers", "nature's basket", "dmart",
      "reliance fresh", "more retail", "grocery", "supermarket",
      "amazon@rapl", "amazon@yapl", "jiomart"],
     "Expense", "Groceries"),

    # ── MEDICAL ───────────────────────────────────────────────────────────────
    (["neuroleap", "care health ins", "oriental insurance co",
      "hospital", "clinic", "pharmacy", "medplus", "apollo",
      "lilavati", "kokilaben", "hinduja", "breach candy",
      "medicine", "doctor", "medical", "health"],
     "Expense", "Medical"),

    # ── CLOTHES ───────────────────────────────────────────────────────────────
    (["vps/zara", "vps/tumi", "vps/sun glass",
      "zara", "h&m", "gap", "myntra", "ajio",
      "clothes", "clothing", "tailoring"],
     "Expense", "Clothes"),

    # ── GIFTS ─────────────────────────────────────────────────────────────────
    (["vps/hamleys", "vps/kitab khana", "hamleys", "kitab khana",
      "cococart", "gift", "present"],
     "Expense", "Gifts"),

    # ── ALCOHOL ───────────────────────────────────────────────────────────────
    (["wine", "beer", "liquor", "alcohol", "spirits", "whisky",
      "hilife", "wb's", "living liquidz", "tasmac"],
     "Expense", "Alcohol"),

    # ── MALHAR ────────────────────────────────────────────────────────────────
    (["malhar", "satoshi", "farm", "m2mferry", "go green nu", "panvel"],
     "Expense", "Malhar"),

    # ── ELECTRICITY & GAS ─────────────────────────────────────────────────────
    (["tata power", "msedcl", "best electricity", "mahanagar gas",
      "mgl", "electricity", "gas bill"],
     "Expense", "Electricity & Gas"),

    # ── HOME OFFICE ───────────────────────────────────────────────────────────
    (["home office", "office supply", "printer", "laptop", "stationery"],
     "Expense", "Home office"),

    # ── ONE TIME CHARGE ───────────────────────────────────────────────────────
    (["vps/ikea india", "ikea"],
     "Expense", "One Time Charge"),

    # ── MAINTENANCE / HOME REPAIR ─────────────────────────────────────────────
    (["plumber", "carpenter", "electrician", "pest control",
      "painting", "repair", "maintenance", "ac service", "deep clean"],
     "Expense", "Maintenance Expense"),

    # ── CASH ATM ──────────────────────────────────────────────────────────────
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
