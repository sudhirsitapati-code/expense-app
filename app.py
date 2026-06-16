"""
app.py
Flask app — all routes: WhatsApp webhook, HTML screens, JSON APIs.
"""

import json
import os
import threading
from datetime import datetime

from apscheduler.schedulers.background import BackgroundScheduler
from dotenv import load_dotenv
from flask import Flask, request, render_template, jsonify, session, redirect, url_for

from src.approval_engine import ApprovalEngine, ExpenseRequest
from src.reconcile import run_reconciliation
from src.acc27_writer import sync_approved_to_history, export_monthly_excel
from src.icici_statement_parser import fetch_and_parse_statements
from src.master_ledger import (
    load_ledger, get_uncertain, update_transaction,
    sync_from_gmail as ledger_sync_gmail,
    get_cc_balance, reconcile_with_approvals,
    import_from_icici_transactions,
    repair_pdf_descriptions, deduplicate_ledger,
    LEDGER_PATH, _load_json as _ml_load_json, _save_json as _ml_save_json,
    _parse_date as _ml_parse_date,
)
from src.whatsapp_handler import (
    build_twiml_reply, parse_incoming,
    send_approval_request, send_approval_result,
    send_auto_approval_notice, send_clarification_request,
    SUDHIR, HOUSEHOLD_MEMBERS,
)
from src import db

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-me")

engine = ApprovalEngine()

BASE_DIR = os.path.dirname(__file__)
DATA_DIR = os.path.join(BASE_DIR, "data")
CONFIG_DIR = os.path.join(BASE_DIR, "config")

# Legacy file paths kept for local dev fallback; production uses db module
APPROVAL_LOG = os.path.join(DATA_DIR, "approval_log.json")
RECONCILE_LOG = os.path.join(DATA_DIR, "reconcile_log.json")
TRANSACTIONS_PATH = os.path.join(DATA_DIR, "icici_transactions.json")

db.init_db()

PENDING_CLARIFICATION: dict = {}
NUMBER_TO_NAME = {v: k for k, v in HOUSEHOLD_MEMBERS.items() if v}


PASSWORDS = {
    "vincent": os.getenv("VINCENT_PASSWORD", "vincent123"),
    "sudhir":  os.getenv("SUDHIR_PASSWORD",  "sudhir123"),
    "ketki":   os.getenv("KETKI_PASSWORD",   "ketki123"),
    "santosh": os.getenv("SANTOSH_PASSWORD", "santosh123"),
}


def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


# ── Helpers ──────────────────────────────────────────────────────────────────

_PATH_TO_KEY = {
    APPROVAL_LOG:    "approval_log",
    RECONCILE_LOG:   "reconcile_log",
    TRANSACTIONS_PATH: "icici_transactions",
}

def _load_json(path):
    key = _PATH_TO_KEY.get(path)
    if key:
        return db.load(key)
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
        db.save(key, data)
        return
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(data, f, indent=2)


def _update_log_entry(request_id: str, updates: dict):
    log = db.load("approval_log")
    for entry in log:
        if entry.get("request_id") == request_id:
            entry.update(updates)
            break
    db.save("approval_log", log)


# ── AUTH ─────────────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = request.form.get("user", "").lower()
        password = request.form.get("password", "")
        if PASSWORDS.get(user) == password:
            session["user"] = user
            return redirect(url_for("index"))
        return render_template("login.html", error="Incorrect name or password.")
    return render_template("login.html", error=None)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ── HTML SCREENS ─────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def index():
    return render_template("index.html", user=session["user"])


# Keep old routes redirecting to /
@app.route("/entry")
@app.route("/dashboard")
@app.route("/report")
def redirect_old():
    return redirect(url_for("index"))


# ── JSON APIs ─────────────────────────────────────────────────────────────────

@app.route("/api/expenses", methods=["GET"])
@login_required
def api_expenses():
    return jsonify(_load_json(APPROVAL_LOG))


@app.route("/api/transactions", methods=["GET"])
def api_transactions():
    return jsonify(_load_json(TRANSACTIONS_PATH))


@app.route("/api/reconcile-log", methods=["GET"])
def api_reconcile_log():
    return jsonify(_load_json(RECONCILE_LOG))


@app.route("/api/mis", methods=["GET"])
@login_required
def api_mis():
    """Return MIS data grouped by super-category.
    Query param: period = month | quarter | ytd  (default: month)

    FY26 column always shows the full-year FY26 actual (from fy26_actuals.json).
    FY27 Budget:
      - month/quarter: scaled to period months
      - ytd: full-year annual budget
    FY27 Actual: what's been approved in the period (month/quarter/ytd).
    """
    period = request.args.get("period", "month")

    with open(os.path.join(CONFIG_DIR, "budget_fy27.json")) as f:
        _bfile = json.load(f)
    budget_annual = _bfile["annual"]   # FY27 annual by ACC26 heading (Blueprint)
    budget_monthly_app = _bfile["monthly"]  # monthly by app category (approval engine)

    # ── FY26 monthly data — from Summaryexpenses sheet, ACC26ver5_MASTER.xlsx ──
    # Calendar month abbreviations: Apr=FY1, May=FY2 … Mar=FY12
    FY26_MONTHLY = {
        "Misc":           {"Apr":16685,"May":836,"Jun":16239,"Jul":6208,"Aug":59760,"Sep":625,"Oct":34168,"Nov":16938,"Dec":2676,"Jan":6562,"Feb":16537,"Mar":11382},
        "Clothes":        {"Apr":0,"May":25990,"Jun":54219,"Jul":66865,"Aug":64337,"Sep":0,"Oct":57982,"Nov":38164,"Dec":19461,"Jan":8290,"Feb":47896,"Mar":7652},
        "Gifts":          {"Apr":6500,"May":3950,"Jun":0,"Jul":0,"Aug":8825,"Sep":0,"Oct":40000,"Nov":33099,"Dec":0,"Jan":3540,"Feb":1571,"Mar":2863},
        "Cash":           {"Apr":10000,"May":0,"Jun":20000,"Jul":25000,"Aug":50000,"Sep":10000,"Oct":10000,"Nov":80000,"Dec":0,"Jan":11600,"Feb":31000,"Mar":10000},
        "Maintenance Expense": {"Apr":17440,"May":9814,"Jun":12775,"Jul":2411,"Aug":3029,"Sep":750,"Oct":22850,"Nov":5887,"Dec":0,"Jan":2500,"Feb":6243,"Mar":4404},
        "Malhar":         {"Apr":40800,"May":53930,"Jun":93923,"Jul":89728,"Aug":51314,"Sep":91350,"Oct":177919,"Nov":165493,"Dec":0,"Jan":81581,"Feb":88880,"Mar":30382},
        "Home office":    {"Apr":88350,"May":25461,"Jun":10450,"Jul":23057,"Aug":24447,"Sep":11933,"Oct":9455,"Nov":24783,"Dec":0,"Jan":7430,"Feb":3096,"Mar":5320},
        "Electricity & Gas": {"Apr":50780,"May":52723,"Jun":52068,"Jul":31128,"Aug":25843,"Sep":24863,"Oct":47174,"Nov":38834,"Dec":0,"Jan":37325,"Feb":24186,"Mar":7616},
        "Alcohol":        {"Apr":14900,"May":0,"Jun":0,"Jul":9400,"Aug":1450,"Sep":0,"Oct":0,"Nov":48700,"Dec":0,"Jan":0,"Feb":0,"Mar":500},
        "Medical":        {"Apr":14143,"May":96808,"Jun":27382,"Jul":77477,"Aug":56988,"Sep":77126,"Oct":72143,"Nov":45345,"Dec":1210,"Jan":57526,"Feb":20200,"Mar":45100},
        "Holiday":        {"Apr":17031,"May":486489,"Jun":132378,"Jul":952200,"Aug":73464,"Sep":4218,"Oct":93531,"Nov":149259,"Dec":0,"Jan":8128,"Feb":56366,"Mar":215510},
        "Groceries":      {"Apr":89175,"May":120368,"Jun":90011,"Jul":106208,"Aug":135331,"Sep":172407,"Oct":155555,"Nov":253751,"Dec":5404,"Jan":220193,"Feb":140067,"Mar":200420},
        "Eating Out":     {"Apr":46057,"May":27258,"Jun":25723,"Jul":48862,"Aug":49917,"Sep":51657,"Oct":69146,"Nov":48007,"Dec":20773,"Jan":53623,"Feb":36546,"Mar":42347},
        "Amma":           {"Apr":27679,"May":16458,"Jun":7276,"Jul":10068,"Aug":12826,"Sep":52229,"Oct":1375,"Nov":-1356,"Dec":0,"Jan":15595,"Feb":-26776,"Mar":-1823},
        "Wellness":       {"Apr":63157,"May":22672,"Jun":38199,"Jul":46159,"Aug":13292,"Sep":59132,"Oct":52176,"Nov":61175,"Dec":0,"Jan":7745,"Feb":13545,"Mar":10649},
        "Ketki":          {"Apr":426656,"May":115151,"Jun":-15927,"Jul":78940,"Aug":126146,"Sep":50147,"Oct":326622,"Nov":203580,"Dec":65053,"Jan":40084,"Feb":27653,"Mar":101024},
        "Staff Salary":   {"Apr":118522,"May":258292,"Jun":269854,"Jul":195408,"Aug":178854,"Sep":176859,"Oct":188398,"Nov":419898,"Dec":99438,"Jan":279669,"Feb":224518,"Mar":221018},
        "Financial Expense / OD Interest": {"Apr":0,"May":0,"Jun":0,"Jul":0,"Aug":0,"Sep":0,"Oct":0,"Nov":0,"Dec":0,"Jan":0,"Feb":0,"Mar":360000},
        "Entertainment":  {"Apr":1148,"May":5278,"Jun":68503,"Jul":40408,"Aug":14280,"Sep":21692,"Oct":19600,"Nov":15839,"Dec":9485,"Jan":4214,"Feb":649,"Mar":25143},
        "One Time Charge":{"Apr":45325,"May":12530,"Jun":4750,"Jul":79984,"Aug":76574,"Sep":149914,"Oct":64441,"Nov":199349,"Dec":71035,"Jan":24050,"Feb":31250,"Mar":11850},
        "Children Education": {"Apr":981577,"May":68350,"Jun":550856,"Jul":101149,"Aug":124383,"Sep":54100,"Oct":38700,"Nov":514160,"Dec":0,"Jan":29700,"Feb":45500,"Mar":524550},
        "Kalpataru Maintenance": {"Apr":35997,"May":35933,"Jun":35879,"Jul":-7041,"Aug":197231,"Sep":36587,"Oct":35879,"Nov":71168,"Dec":0,"Jan":35879,"Feb":35879,"Mar":35933},
        "Charity":        {"Apr":20000,"May":0,"Jun":0,"Jul":167800,"Aug":105725,"Sep":24048,"Oct":300000,"Nov":115000,"Dec":150000,"Jan":625100,"Feb":0,"Mar":18000},
        "Uspaar":         {"Apr":133783,"May":140978,"Jun":214372,"Jul":97910,"Aug":204600,"Sep":72828,"Oct":38984,"Nov":170003,"Dec":0,"Jan":103660,"Feb":40350,"Mar":86325},
        "Insurance":      {"Apr":118000,"May":0,"Jun":0,"Jul":0,"Aug":0,"Sep":0,"Oct":102820,"Nov":0,"Dec":142800,"Jan":0,"Feb":0,"Mar":0},
        "Home Loan":      {"Apr":903259,"May":1146721,"Jun":916551,"Jul":1201557,"Aug":554547,"Sep":992227,"Oct":963511,"Nov":1007410,"Dec":1777170,"Jan":1389117,"Feb":928825,"Mar":1285719},
        "Tax":            {"Apr":0,"May":0,"Jun":0,"Jul":13519347,"Aug":0,"Sep":3630000,"Oct":61499379,"Nov":0,"Dec":0,"Jan":0,"Feb":3848950,"Mar":0},
    }
    # Calendar month number → FY month abbreviation
    CAL_TO_FY_MON = {4:"Apr",5:"May",6:"Jun",7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec",1:"Jan",2:"Feb",3:"Mar"}

    log = _load_json(APPROVAL_LOG)

    # ── Period months ─────────────────────────────────────────────────────────
    now = datetime.now()
    cal_month = now.month
    fy_start_year = now.year if cal_month >= 4 else now.year - 1

    def _period_months(p):
        if p == "month":
            return [now.strftime("%Y-%m")]
        elif p == "quarter":
            fy_month = (cal_month - 4) % 12 + 1
            q_start_fy = ((fy_month - 1) // 3) * 3 + 1
            months = []
            for offset in range(3):
                cm_idx = (q_start_fy - 1 + offset) % 12
                cal = (cm_idx + 4 - 1) % 12 + 1
                yr = fy_start_year if cal >= 4 else fy_start_year + 1
                months.append(f"{yr}-{cal:02d}")
            return months
        else:  # ytd
            months, cm, yr = [], 4, fy_start_year
            while True:
                months.append(f"{yr}-{cm:02d}")
                if yr == now.year and cm == cal_month:
                    break
                cm += 1
                if cm == 13:
                    cm, yr = 1, yr + 1
            return months

    period_months = _period_months(period)
    n_months = len(period_months)

    # FY month names covered by this period (for FY26 lookup)
    fy_mon_names = [CAL_TO_FY_MON[int(m.split("-")[1])] for m in period_months]

    # ── Heading → super-category ──────────────────────────────────────────────
    HEADING_SUPER = {
        "Groceries":"Household","Staff Salary":"Household","Electricity & Gas":"Household",
        "Misc":"Household","Cash":"Household",
        "Alcohol":"Personal","Wellness":"Personal",
        "Clothes":"Family","Gifts":"Family","Gift":"Family","Medical":"Family",
        "Amma":"Family","Ketki":"Family","Children Education":"Family",
        "Charity":"Giving","Uspaar":"Giving",
        "Holiday":"Lifestyle","Eating Out":"Lifestyle","Entertainment":"Lifestyle",
        "Malhar":"Property","Maintenance Expense":"Property","Home office":"Property",
        "One Time Charge":"Property","Kalpataru Maintenance":"Property",
        "Financial Expense":"Financial","Financial Expense / OD Interest":"Financial",
        "Insurance":"Financial","Home Loan":"Financial","Tax":"Financial",
    }
    SUPER_ORDER = ["Household","Personal","Family","Giving","Lifestyle","Property","Financial"]

    # App category → ACC26 heading (for FY27 actual from approval log)
    APP_TO_HEADING = {
        "groceries":"Groceries","staff":"Staff Salary","utilities":"Electricity & Gas",
        "miscellaneous":"Misc","personal_care":"Wellness","clothing":"Clothes",
        "gifts":"Gifts","medical":"Medical","education":"Children Education",
        "dining":"Eating Out","entertainment":"Entertainment","transport":"Holiday",
        "maintenance":"Maintenance Expense","home_repair":"One Time Charge",
    }

    # ── FY26 for period: sum actual monthly data for the matched FY months ────
    def _fy26_period(heading):
        monthly = FY26_MONTHLY.get(heading, {})
        return max(0, sum(monthly.get(m, 0) for m in fy_mon_names))

    def _fy26_full_year(heading):
        monthly = FY26_MONTHLY.get(heading, {})
        return max(0, sum(monthly.values()))

    # ── FY27 budget scaled to period ─────────────────────────────────────────
    def _budget(heading):
        annual = budget_annual.get(heading, 0)
        return annual if period == "ytd" else round(annual / 12 * n_months)

    # ── FY27 actual from master ledger ────────────────────────────────────────
    from src.master_ledger import _parse_date as _ml_parse_date
    fy27_actual: dict = {}
    _fy27_start = datetime(2026, 4, 1)
    for txn in load_ledger():
        if (txn.get("type") or "").lower() not in ("expense", "official"):
            continue
        if txn.get("uncertain"):
            continue
        dt = _ml_parse_date(txn.get("date", ""))
        if not dt or dt < _fy27_start:
            continue
        ym = dt.strftime("%Y-%m")
        if ym not in period_months:
            continue
        heading = txn.get("heading", "Misc") or "Misc"
        amt = float(txn.get("debit", 0) or 0)
        if amt > 0:
            fy27_actual[heading] = fy27_actual.get(heading, 0) + amt

    # ── Build grouped rows ────────────────────────────────────────────────────
    all_headings = set(budget_annual.keys()) | set(FY26_MONTHLY.keys()) | set(fy27_actual.keys())
    by_super: dict = {s: [] for s in SUPER_ORDER}

    for heading in sorted(all_headings):
        super_cat = HEADING_SUPER.get(heading, "Household")
        fy26 = round(_fy26_period(heading))
        budget = _budget(heading)
        actual = round(fy27_actual.get(heading, 0))
        pct = round(actual / budget * 100) if budget else 0
        row = {
            "category": heading,
            "fy26_actual": fy26,
            "fy27_budget": budget,
            "fy27_actual": actual,
            "pct": pct,
        }
        if period == "ytd":
            row["fy26_full_year"] = round(_fy26_full_year(heading))
        by_super.setdefault(super_cat, []).append(row)

    groups = []
    grand = {"fy26": 0, "fy26_full_year": 0, "budget": 0, "actual": 0}
    for super_cat in SUPER_ORDER:
        rows = by_super.get(super_cat, [])
        if not rows:
            continue
        sub = {
            "fy26":   sum(r["fy26_actual"] for r in rows),
            "budget": sum(r["fy27_budget"] for r in rows),
            "actual": sum(r["fy27_actual"] for r in rows),
        }
        if period == "ytd":
            sub["fy26_full_year"] = sum(r.get("fy26_full_year", 0) for r in rows)
            grand["fy26_full_year"] += sub["fy26_full_year"]
        sub["pct"] = round(sub["actual"] / sub["budget"] * 100) if sub["budget"] else 0
        grand["fy26"]   += sub["fy26"]
        grand["budget"] += sub["budget"]
        grand["actual"] += sub["actual"]
        groups.append({"super_category": super_cat, "rows": rows, "subtotal": sub})

    grand["pct"] = round(grand["actual"] / grand["budget"] * 100) if grand["budget"] else 0

    return jsonify({
        "period": period,
        "period_months": period_months,
        "groups": groups,
        "grand": grand,
    })


@app.route("/api/cash-recon", methods=["GET"])
def api_cash_recon():
    """ATM withdrawals vs cash-authorised payments."""
    transactions = _load_json(TRANSACTIONS_PATH)
    log = _load_json(APPROVAL_LOG)

    # ATM withdrawals from ICICI transactions
    atm_txns = [t for t in transactions if "atm" in (t.get("description") or "").lower() and t.get("type") == "debit"]
    atm_total = sum(t.get("amount", 0) for t in atm_txns)

    # Cash payments from approval log (this month)
    month_prefix = datetime.now().strftime("%Y-%m")
    cash_approved = [
        e for e in log
        if e.get("payment_method") == "cash"
        and e.get("action") in ("AUTO_APPROVE", "APPROVED", "APPROVED_LOWER")
        and (e.get("timestamp") or "").startswith(month_prefix)
    ]
    cash_paid_total = sum((e.get("approved_amount") or e.get("amount", 0)) for e in cash_approved)

    # Monthly breakdown
    monthly: dict = {}
    for t in atm_txns:
        m = (t.get("date") or "")[:7]
        if m:
            monthly.setdefault(m, {"atm": 0, "cash_paid": 0})
            monthly[m]["atm"] += t.get("amount", 0)

    for e in [x for x in log if x.get("payment_method") == "cash" and x.get("action") in ("AUTO_APPROVE", "APPROVED", "APPROVED_LOWER")]:
        m = (e.get("timestamp") or "")[:7]
        if m:
            monthly.setdefault(m, {"atm": 0, "cash_paid": 0})
            monthly[m]["cash_paid"] += e.get("approved_amount") or e.get("amount", 0)

    monthly_rows = [{"month": k, **v} for k, v in sorted(monthly.items(), reverse=True)]

    return jsonify({
        "atm_total": atm_total,
        "cash_paid_total": cash_paid_total,
        "monthly_rows": monthly_rows,
    })


@app.route("/api/submit-expense", methods=["POST"])
def api_submit_expense():
    """Handle expense submission from the web form."""
    data = request.get_json()

    # Follow-up answer to a pending clarification
    if "request_id" in data and "followup_answer" in data:
        req_id = data["request_id"]
        answer = data["followup_answer"]
        req = PENDING_CLARIFICATION.get(req_id)
        if req:
            req.description = f"{req.description} [{answer}]"
            PENDING_CLARIFICATION.pop(req_id, None)

            if data.get("followup_round", 1) >= 2:
                # Max rounds hit — escalate with unclear flag
                from src.approval_engine import ApprovalDecision
                decision = ApprovalDecision(
                    request_id=req_id, action="ESCALATE",
                    reason="Unclear pricing after 2 follow-up rounds"
                )
                decision.escalation_message = engine._build_escalation_message(req, "\n⚠️ Pricing unclear after follow-up")
                send_approval_request(decision.escalation_message)
                return jsonify({"action": "ESCALATE", "request_id": req_id,
                                "vendor": req.vendor, "amount": req.amount})

            decision = engine.evaluate(req)
        else:
            return jsonify({"action": "ESCALATE", "request_id": req_id})
    else:
        # Fresh submission
        try:
            req = ExpenseRequest(
                submitter=data["submitter"],
                vendor=data["vendor"],
                amount=float(data["amount"]),
                category=data["category"],
                description=data["description"],
                payment_method=data.get("payment_method", "upi"),
                is_post_facto=data.get("is_post_facto", False),
            )
        except (KeyError, ValueError) as e:
            return jsonify({"error": str(e)}), 400

        decision = engine.evaluate(req)

    if decision.action == "PENDING_CLARIFICATION":
        PENDING_CLARIFICATION[decision.request_id] = req

    if decision.action == "AUTO_APPROVE":
        sync_approved_to_history()
        send_auto_approval_notice(req.submitter, req.vendor, req.amount, decision.request_id)
    elif decision.action == "ESCALATE" and decision.escalation_message:
        send_approval_request(decision.escalation_message)

    return jsonify({
        "action": decision.action,
        "request_id": decision.request_id,
        "vendor": req.vendor,
        "amount": req.amount,
        "market_rate": decision.market_rate,
        "market_status": decision.market_status,
        "budget_alert": decision.budget_alert,
        "follow_up_question": decision.follow_up_question,
        "follow_up_options": decision.follow_up_options,
    })


@app.route("/api/decide", methods=["POST"])
def api_decide():
    """Approve or reject from dashboard buttons."""
    data = request.get_json()
    request_id = data.get("request_id")
    response = data.get("response", "")

    log = _load_json(APPROVAL_LOG)
    entry = next((e for e in log if e.get("request_id") == request_id), None)
    if not entry:
        return jsonify({"error": "not found"}), 404

    engine.update_log_with_sudhir_response(request_id, response)

    if response.upper() == "Y":
        send_approval_result(entry["submitter"], entry["vendor"], entry["amount"], approved=True, request_id=request_id)
        sync_approved_to_history()
    elif response.upper() == "N":
        send_approval_result(entry["submitter"], entry["vendor"], entry["amount"], approved=False, request_id=request_id)

    return jsonify({"status": "ok"})


def _approval_to_ledger_entry(e: dict) -> dict:
    """Convert an approval log entry into a master ledger transaction."""
    import hashlib
    from src.master_ledger import _fy_info, _parse_date, classify_transaction
    paid_at = e.get("confirmed_at") or e.get("response_timestamp") or e.get("timestamp","")
    date_str = paid_at[:10]
    dt = _parse_date(date_str) or datetime.now()
    fy = _fy_info(dt)
    amount = float(e.get("approved_amount") or e.get("amount") or 0)
    raw = f"{date_str}|approval|{e.get('vendor','')}|{amount:.2f}"
    txn_id = hashlib.sha1(raw.encode()).hexdigest()[:16]

    APP_TO_HEADING = {
        "groceries":"Groceries","staff":"Staff Salary","utilities":"Electricity & Gas",
        "miscellaneous":"Misc","personal_care":"Wellness","clothing":"Clothes",
        "gifts":"Gifts","medical":"Medical","education":"Children Education",
        "dining":"Eating Out","entertainment":"Entertainment","transport":"Holiday",
        "maintenance":"Maintenance Expense","home_repair":"One Time Charge",
    }
    heading = APP_TO_HEADING.get(e.get("category",""), "Misc")

    txn = {
        "txn_id":          txn_id,
        "date":            date_str,
        "fy_month_no":     fy["fy_month_no"],
        "fy_month_name":   fy["fy_month_name"],
        "fy_year":         fy["fy_year"],
        "account":         "cash/upi",
        "account_type":    e.get("payment_method","cash"),
        "bank":            "approval",
        "raw_description": e.get("description",""),
        "paid_to":         e.get("vendor",""),
        "debit":           amount,
        "credit":          0,
        "type":            "expense",
        "heading":         heading,
        "remarks":         f"Approved by Sudhir. Ref: {e.get('request_id','')}",
        "uncertain":       False,
        "uncertain_fields":[],
        "confidence":      "approval",
        "source":          "approval_log",
        "gmail_id":        None,
        "ai_saving_tip":   None,
        "saving_agreed":   None,
        "reconciled_with": e.get("request_id"),
        "created_at":      datetime.now().isoformat(),
    }
    return txn


def _sync_approvals_to_ledger():
    """Add all confirmed-paid approval entries to master ledger (idempotent)."""
    from src.master_ledger import load_ledger
    log    = db.load("approval_log")
    ledger = db.load("master_ledger")
    existing_ids = {t["txn_id"] for t in ledger}

    added = 0
    for e in log:
        if not e.get("confirmed_paid"):
            continue
        if e.get("action") not in ("AUTO_APPROVE","APPROVED","APPROVED_LOWER"):
            continue
        txn = _approval_to_ledger_entry(e)
        if txn["txn_id"] in existing_ids:
            continue
        ledger.append(txn)
        existing_ids.add(txn["txn_id"])
        added += 1

    if added:
        ledger.sort(key=lambda t: t.get("date",""), reverse=True)
        db.save("master_ledger", ledger)
    return added


@app.route("/api/mark-paid", methods=["POST"])
def api_mark_paid():
    """Mark a cash expense as confirmed paid and push it into the master ledger."""
    data = request.get_json()
    request_id = data.get("request_id")
    now_iso = datetime.now().isoformat()
    _update_log_entry(request_id, {"confirmed_paid": True, "confirmed_at": now_iso})

    # Push this entry straight into the ledger
    log   = db.load("approval_log")
    entry = next((e for e in log if e.get("request_id") == request_id), None)
    if entry:
        ledger = db.load("master_ledger")
        txn = _approval_to_ledger_entry(entry)
        if not any(t["txn_id"] == txn["txn_id"] for t in ledger):
            ledger.insert(0, txn)
            db.save("master_ledger", ledger)

    return jsonify({"status": "ok"})


@app.route("/api/sync-approvals-to-ledger", methods=["POST"])
@login_required
def api_sync_approvals_to_ledger():
    """Backfill all confirmed-paid approvals into the master ledger."""
    added = _sync_approvals_to_ledger()
    return jsonify({"status": "ok", "added": added})


@app.route("/api/ignore-unauth", methods=["POST"])
def api_ignore_unauth():
    """Mark an unmatched bank debit as known/ignored."""
    data = request.get_json()
    gmail_id = data.get("gmail_id")
    recon = _load_json(RECONCILE_LOG)
    for entry in recon:
        if entry.get("gmail_id") == gmail_id:
            entry["ignored"] = True
            entry["ignored_at"] = datetime.now().isoformat()
            break
    _save_json(RECONCILE_LOG, recon)
    return jsonify({"status": "ok"})


@app.route("/api/sync-statements", methods=["POST"])
@login_required
def api_sync_statements():
    try:
        result = fetch_and_parse_statements()
        return jsonify({"status": "ok", **result})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/transactions/<txn_id>", methods=["PATCH"])
@login_required
def api_update_transaction(txn_id):
    """Update editable fields of an ICICI transaction (paid_to, acc_type, heading, remarks)."""
    data = request.get_json()
    allowed = {"paid_to", "acc_type", "heading", "remarks"}
    updates = {k: v for k, v in data.items() if k in allowed}
    if not updates:
        return jsonify({"error": "no valid fields"}), 400
    transactions = _load_json(TRANSACTIONS_PATH)
    for t in transactions:
        if t.get("txn_id") == txn_id:
            t.update(updates)
            # Mark as manually verified once user edits classification
            if "acc_type" in updates or "heading" in updates:
                t["confidence"] = "manual"
            _save_json(TRANSACTIONS_PATH, transactions)
            return jsonify({"status": "ok"})
    return jsonify({"error": "not found"}), 404


# ── MASTER LEDGER APIs ────────────────────────────────────────────────────────

@app.route("/api/master-ledger", methods=["GET"])
@login_required
def api_master_ledger():
    """Return master ledger entries. Filters: uncertain, account, bank, type, heading, month, q."""
    txns = load_ledger()

    # Optional filters
    only_uncertain = request.args.get("uncertain") == "1"
    account_filter = request.args.get("account")
    bank_filter    = request.args.get("bank")
    type_filter    = request.args.get("type")
    heading_filter = request.args.get("heading")
    month_filter   = request.args.get("month")   # YYYY-MM
    search         = request.args.get("q","").lower()

    if only_uncertain:
        txns = [t for t in txns if t.get("uncertain")]
    if account_filter:
        txns = [t for t in txns if t.get("account") == account_filter]
    if bank_filter:
        txns = [t for t in txns if t.get("bank") == bank_filter]
    if type_filter:
        txns = [t for t in txns if (t.get("type") or "").lower() == type_filter.lower()]
    if heading_filter:
        txns = [t for t in txns if (t.get("heading") or "") == heading_filter]
    if month_filter:
        txns = [t for t in txns if (t.get("date",""))[:7] == month_filter
                or t.get("date","").endswith(month_filter[-2:] + "/" + month_filter[:4])]
    if search:
        txns = [t for t in txns
                if search in (t.get("raw_description","")).lower()
                or search in (t.get("paid_to","")).lower()
                or search in (t.get("heading") or "").lower()]

    # Summary stats
    total_debit  = sum(t.get("debit",0)  for t in txns)
    total_credit = sum(t.get("credit",0) for t in txns)
    uncertain_ct = sum(1 for t in txns if t.get("uncertain"))

    # Accounts list for filter dropdown
    all_accounts = sorted({t.get("account","") for t in load_ledger() if t.get("account")})

    return jsonify({
        "transactions": txns,
        "total_debit":  total_debit,
        "total_credit": total_credit,
        "uncertain":    uncertain_ct,
        "accounts":     all_accounts,
        "count":        len(txns),
    })


@app.route("/api/master-ledger/sync", methods=["POST"])
@login_required
def api_ledger_sync():
    """Pull new bank alert emails from Gmail only (fast). PDF statements use /sync-statements."""
    data  = request.get_json() or {}
    days  = int(data.get("days", 90))
    force = bool(data.get("force", False))
    result = ledger_sync_gmail(days_back=days, force=force)

    # Fix account names, reclassify unknown entries, remove cross-source duplicates
    repair_pdf_descriptions()
    result["deduped"] = deduplicate_ledger()

    # Reconcile with approval log
    log = _load_json(APPROVAL_LOG)
    matched = reconcile_with_approvals(log)
    result["reconciled"] = matched
    return jsonify({"status": "ok", **result})


_stmt_sync_state = {"running": False, "result": None, "error": None, "started_at": None}


def _run_statement_sync(force: bool):
    """Heavy PDF sync — runs in background thread to avoid request timeout/OOM."""
    result = {}
    try:
        # Purge pre-FY27 pdf_import entries on force
        if force:
            _fy27_start = datetime(2026, 4, 1)
            _ledger = _ml_load_json(LEDGER_PATH)
            _before = len(_ledger)
            _ledger = [
                t for t in _ledger
                if not (
                    t.get("source") == "pdf_import"
                    and (lambda d: d is None or d < _fy27_start)(_ml_parse_date(t.get("date", "")))
                )
            ]
            if len(_ledger) < _before:
                _ml_save_json(LEDGER_PATH, _ledger)
                result["pdf_purged"] = _before - len(_ledger)

        pdf_result = fetch_and_parse_statements(force_reprocess=force)
        result["pdf_statements"] = pdf_result.get("statements", 0)
        result["pdf_new"]        = pdf_result.get("new", 0)
        result["pdf_imported"]   = import_from_icici_transactions()
        result["pdf_repaired"]   = repair_pdf_descriptions()
        result["deduped"]        = deduplicate_ledger()
        log = _load_json(APPROVAL_LOG)
        result["reconciled"]     = reconcile_with_approvals(log)
        result["status"] = "ok"
    except Exception as e:
        result["status"] = "error"
        result["error"]  = str(e)
    finally:
        _stmt_sync_state["result"]  = result
        _stmt_sync_state["running"] = False


@app.route("/api/master-ledger/sync-statements", methods=["POST"])
@login_required
def api_ledger_sync_statements():
    """Start PDF statement sync in background. Returns immediately — poll /sync-status."""
    if _stmt_sync_state["running"]:
        return jsonify({"status": "already_running", "message": "Sync already in progress — check /api/master-ledger/sync-status"})
    data  = request.get_json() or {}
    force = bool(data.get("force", False))
    _stmt_sync_state["running"]    = True
    _stmt_sync_state["result"]     = None
    _stmt_sync_state["error"]      = None
    _stmt_sync_state["started_at"] = datetime.now().isoformat()
    threading.Thread(target=_run_statement_sync, args=(force,), daemon=True).start()
    return jsonify({"status": "started", "message": "Sync running in background — poll /api/master-ledger/sync-status"})


@app.route("/api/master-ledger/sync-status", methods=["GET"])
@login_required
def api_ledger_sync_status():
    """Check status of background PDF statement sync."""
    return jsonify({
        "running":    _stmt_sync_state["running"],
        "started_at": _stmt_sync_state["started_at"],
        "result":     _stmt_sync_state["result"],
    })


@app.route("/api/master-ledger/fix-7281", methods=["POST"])
@login_required
def api_fix_7281():
    """Directly reclassify all non-manual ICICI-7281 entries by description keywords."""
    INTEREST_KW = ["int.coll", "int coll", "renewal", "sgst", "cgst", "interest", "bank charge", "processing fee"]
    TRANSFER_KW = ["inft", "neft", "imps", "rtgs", "upi", "transfer", "self", "inf/"]

    ledger = _ml_load_json(LEDGER_PATH)
    fixed = 0
    for txn in ledger:
        if txn.get("account") != "ICICI-7281":
            continue
        if txn.get("confidence") == "manual":
            continue
        desc = (txn.get("raw_description") or "").lower()

        # Normalise date format while we're here
        if txn.get("date") and "-" in txn.get("date", ""):
            dt = _ml_parse_date(txn["date"])
            if dt:
                txn["date"] = dt.strftime("%d/%m/%Y")

        if any(k in desc for k in INTEREST_KW):
            txn["type"] = "Expense"
            txn["heading"] = "Financial Expense"
            txn["uncertain"] = False
            fixed += 1
        elif any(k in desc for k in TRANSFER_KW) or not desc or txn.get("heading") in (None, "", "null"):
            txn["type"] = "Transfer"
            txn["heading"] = "Interbank"
            txn["uncertain"] = not desc
            fixed += 1

    if fixed:
        _ml_save_json(LEDGER_PATH, ledger)
    return jsonify({"status": "ok", "fixed": fixed})


@app.route("/api/master-ledger/bulk-classify", methods=["POST"])
@login_required
def api_bulk_classify():
    """Bulk-set type+heading for transactions matching account + date range.
    Skips transfers, income, investments, and manually-confirmed entries.
    Body: {accounts, date_from, date_to, type, heading}
    """
    data       = request.get_json() or {}
    accounts   = [a.upper() for a in data.get("accounts", [])]
    date_from  = _ml_parse_date(data.get("date_from", ""))
    date_to    = _ml_parse_date(data.get("date_to", ""))
    new_type   = data.get("type", "Expense")
    new_heading= data.get("heading", "")

    if not accounts or not date_from or not date_to or not new_heading:
        return jsonify({"error": "accounts, date_from, date_to, heading required"}), 400

    SKIP_TYPES    = {"transfer", "income", "investment", "Transfer", "Income", "Investment"}
    SKIP_HEADINGS = {"Salary", "Interbank", "Home Loan", "Loan Repayment", "Tax",
                     "Financial Expense", "Insurance", "GCPL Share Sale", "Foreign Investment",
                     "Uspaar", "Art", "Property Investment"}

    ledger = _ml_load_json(LEDGER_PATH)
    updated = 0
    for txn in ledger:
        if txn.get("account", "").upper() not in accounts:
            continue
        if txn.get("confidence") == "manual":
            continue
        if txn.get("type") in SKIP_TYPES or txn.get("heading") in SKIP_HEADINGS:
            continue
        dt = _ml_parse_date(txn.get("date", ""))
        if not dt or not (date_from <= dt <= date_to):
            continue
        txn["type"]       = new_type
        txn["heading"]    = new_heading
        txn["confidence"] = "manual"
        txn["uncertain"]  = False
        updated += 1

    if updated:
        _ml_save_json(LEDGER_PATH, ledger)
    return jsonify({"status": "ok", "updated": updated})


@app.route("/api/debug/ledger-accounts", methods=["GET"])
@login_required
def api_debug_ledger_accounts():
    """Show distinct account names and sample descriptions per account."""
    from collections import defaultdict
    ledger = _ml_load_json(LEDGER_PATH)
    by_account = defaultdict(list)
    for t in ledger:
        by_account[t.get("account", "?")].append({
            "desc": t.get("raw_description", "")[:80],
            "type": t.get("type", ""),
            "heading": t.get("heading", ""),
            "source": t.get("source", ""),
        })
    return jsonify({
        acct: {"count": len(txns), "samples": txns[:5]}
        for acct, txns in sorted(by_account.items())
    })


@app.route("/api/debug/seq-lookup", methods=["GET"])
@login_required
def api_debug_seq_lookup():
    """Look up transactions by seq numbers. ?seqs=1,2,6,8"""
    seqs = [int(s.strip()) for s in (request.args.get("seqs") or "").split(",") if s.strip().isdigit()]
    if not seqs:
        return jsonify({"error": "provide ?seqs=1,2,3"}), 400
    ledger = load_ledger()
    hits = {t["seq"]: t for t in ledger if t.get("seq") in seqs}
    return jsonify({"results": [
        {k: hits[s].get(k) for k in ("seq","txn_id","date","account","debit","credit","type","heading","source","confidence","raw_description","paid_to")}
        for s in seqs if s in hits
    ]})


@app.route("/api/debug/search-ledger", methods=["GET"])
@login_required
def api_debug_search_ledger():
    """Search ledger by keyword across description, paid_to, account."""
    q = (request.args.get("q") or "").lower()
    if not q:
        return jsonify({"error": "provide ?q=keyword"}), 400
    ledger = load_ledger()
    hits = [t for t in ledger if any(
        q in str(t.get(f) or "").lower()
        for f in ("raw_description", "paid_to", "account", "heading", "type", "description")
    )]
    hits.sort(key=lambda t: t.get("date", ""))
    return jsonify({"count": len(hits), "results": [
        {k: t.get(k) for k in ("txn_id","date","account","debit","credit","type","heading","source","confidence","raw_description","paid_to","description")}
        for t in hits
    ]})


@app.route("/api/debug/mis-actuals", methods=["GET"])
@login_required
def api_debug_mis_actuals():
    """Show what the MIS endpoint aggregates from the master ledger for FY27."""
    from src.master_ledger import _parse_date as _ml_parse_date
    from datetime import datetime
    _fy27_start = datetime(2026, 4, 1)
    fy27_actual = {}
    skipped = []
    for txn in load_ledger():
        dt = _ml_parse_date(txn.get("date", ""))
        if not dt or dt < _fy27_start:
            continue
        reason = None
        if (txn.get("type") or "").lower() not in ("expense", "official"):
            reason = f"type={txn.get('type')}"
        elif txn.get("uncertain"):
            reason = "uncertain=True"
        elif not float(txn.get("debit", 0) or 0):
            reason = "debit=0"
        if reason:
            skipped.append({"seq": txn.get("seq"), "date": txn.get("date"), "heading": txn.get("heading"), "type": txn.get("type"), "uncertain": txn.get("uncertain"), "debit": txn.get("debit"), "reason": reason})
        else:
            h = txn.get("heading") or "Misc"
            fy27_actual[h] = fy27_actual.get(h, 0) + float(txn.get("debit", 0))
    return jsonify({"included": fy27_actual, "skipped_count": len(skipped), "skipped_sample": skipped[:30]})


@app.route("/api/master-ledger/<txn_id>", methods=["PATCH"])
@login_required
def api_ledger_update(txn_id):
    """Update type, heading, paid_to, remarks, saving_agreed for a transaction."""
    data    = request.get_json() or {}
    allowed = {"paid_to","type","heading","remarks","saving_agreed"}
    updates = {k: v for k, v in data.items() if k in allowed}
    if not updates:
        return jsonify({"error": "no valid fields"}), 400
    ok = update_transaction(txn_id, updates)
    return jsonify({"status": "ok"}) if ok else (jsonify({"error": "not found"}), 404)


@app.route("/api/cc-balance", methods=["GET"])
@login_required
def api_cc_balance():
    """Credit card unpaid balances per card."""
    return jsonify(get_cc_balance())


@app.route("/api/approval-log-raw", methods=["GET"])
@login_required
def api_approval_log_raw():
    """Debug: return last 30 raw entries from approval log."""
    log = _load_json(APPROVAL_LOG)
    return jsonify({"total": len(log), "last_30": log[-30:]})


@app.route("/api/db-status", methods=["GET"])
@login_required
def api_db_status():
    """Debug: show DB connection health and stored keys."""
    return jsonify(db.db_status())


@app.route("/api/approvals-structured", methods=["GET"])
@login_required
def api_approvals_structured():
    """Return approval log structured into pending / month / unauthorized / tracker."""
    log   = _load_json(APPROVAL_LOG)
    recon = _load_json(RECONCILE_LOG)

    now          = datetime.now()
    # Allow caller to pass ?month=YYYY-MM; default to current month
    month_prefix = request.args.get("month") or now.strftime("%Y-%m")

    approved_actions = ("AUTO_APPROVE", "APPROVED", "APPROVED_LOWER")

    def _effective_month(e):
        """Use approval date for approved entries; submission date for auto-approvals."""
        if e.get("action") in ("APPROVED", "APPROVED_LOWER"):
            return (e.get("response_timestamp") or e.get("timestamp",""))[:7]
        return (e.get("timestamp",""))[:7]

    today_prefix = now.strftime("%Y-%m")
    # All months that have at least one approved entry (for the month picker), not in the future
    months_with_data = sorted({
        _effective_month(e)
        for e in log
        if (e.get("action") in approved_actions
            or (e.get("action") == "ESCALATE" and "sudhir_response" in e))
        and len(_effective_month(e)) == 7
        and _effective_month(e) <= today_prefix
    }, reverse=True)

    pending      = [e for e in log if e.get("action") == "ESCALATE" and "sudhir_response" not in e]
    this_month   = [e for e in log
                    if _effective_month(e) == month_prefix
                    and (e.get("action") in approved_actions
                         or (e.get("action") == "ESCALATE" and "sudhir_response" in e))]
    unauthorized = [e for e in recon if not e.get("matched") and not e.get("is_recurring") and not e.get("ignored")]
    tracker      = [e for e in log
                    if e.get("action") in approved_actions
                    and not e.get("confirmed_paid")]

    return jsonify({
        "pending":           pending,
        "this_month":        this_month,
        "selected_month":    month_prefix,
        "months_with_data":  months_with_data,
        "unauthorized":      unauthorized,
        "tracker":           tracker,
    })


@app.route("/api/insights", methods=["GET"])
@login_required
def api_insights():
    """AI-generated spending insights from master ledger + approval log."""
    try:
        from openai import AzureOpenAI
        client = AzureOpenAI(
            api_key=os.getenv("AZURE_OPENAI_KEY"),
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
            api_version=os.getenv("AZURE_OPENAI_API_VERSION","2024-12-01-preview"),
        )
        deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT","gpt-5.5")
    except Exception:
        return jsonify({"insights":[]})

    ledger = load_ledger()
    now    = datetime.now()
    month_prefix = now.strftime("%Y-%m")

    # Last 3 months summary by heading
    from collections import defaultdict
    by_heading: dict = defaultdict(float)
    for t in ledger:
        if t.get("debit") and t.get("heading"):
            by_heading[t["heading"]] += t["debit"]

    # Budget from file
    budget_path = os.path.join(CONFIG_DIR, "budget_fy27.json")
    with open(budget_path) as f:
        budget_annual = json.load(f).get("annual",{})

    summary = [
        {"heading": h, "spend": round(v), "budget": budget_annual.get(h,0)}
        for h, v in sorted(by_heading.items(), key=lambda x: -x[1])[:15]
    ]

    prompt = f"""You are a personal finance advisor for an Indian household (Mumbai, upper-income).
Analyse this spending data and give 5 concise, actionable insights to reduce expenses.
Focus on high-spend categories vs budget, patterns worth questioning, and concrete alternatives.

Spending summary (from bank transactions): {json.dumps(summary)}

Reply as JSON array of 5 objects: [{{"title":"...", "detail":"...", "category":"...", "potential_saving":"₹X,XXX/month"}}]
Be specific, not generic. No obvious tips."""

    try:
        resp = client.chat.completions.create(
            model=deployment, max_tokens=800,
            messages=[{"role":"system","content":"Reply only with JSON."},
                      {"role":"user","content":prompt}]
        )
        text = resp.choices[0].message.content.strip()
        if text.startswith("```"):
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        insights = json.loads(text)
    except Exception as e:
        insights = [{"title":"Unable to generate insights","detail":str(e),"category":"","potential_saving":""}]

    return jsonify({"insights": insights})


@app.route("/export", methods=["GET"])
def export():
    month_str = request.args.get("month")
    year, month = (int(x) for x in month_str.split("-")) if month_str else (None, None)
    path = export_monthly_excel(year=year, month=month)
    return jsonify({"status": "exported", "file": path})


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


# ── WHATSAPP WEBHOOK ──────────────────────────────────────────────────────────

def _handle_sudhir_response(body: str):
    parts = body.strip().split()
    command = parts[0].upper()

    log = _load_json(APPROVAL_LOG)
    pending = [e for e in log if e.get("action") == "ESCALATE" and "sudhir_response" not in e]
    if not pending:
        return build_twiml_reply("No pending approvals found.")

    latest = pending[-1]
    request_id = latest["request_id"]
    submitter = latest["submitter"]
    vendor = latest["vendor"]
    amount = latest["amount"]

    engine.update_log_with_sudhir_response(request_id, body.strip())

    if command == "Y":
        send_approval_result(submitter, vendor, amount, approved=True, request_id=request_id)
        sync_approved_to_history()
        return build_twiml_reply(f"✅ Approved. {submitter.title()} notified.\nRef: {request_id}")
    elif command == "N":
        send_approval_result(submitter, vendor, amount, approved=False, request_id=request_id)
        return build_twiml_reply(f"❌ Rejected. {submitter.title()} notified.\nRef: {request_id}")
    elif command == "L" and len(parts) > 1:
        try:
            lower_amount = float(parts[1].replace(",", ""))
            send_approval_result(submitter, vendor, amount, approved=True,
                                 request_id=request_id, approved_amount=lower_amount)
            sync_approved_to_history()
            return build_twiml_reply(f"✅ Approved at Rs {lower_amount:,.0f}.\nRef: {request_id}")
        except ValueError:
            return build_twiml_reply("Invalid amount. Use: L 5000")

    return build_twiml_reply("Use Y / N / L <amount> to respond.")


def _handle_member_message(sender: str, body: str):
    submitter = NUMBER_TO_NAME.get(sender, "unknown")

    # Clarification reply?
    pending_id = next((rid for rid, req in PENDING_CLARIFICATION.items()
                       if req.submitter.lower() == submitter.lower()), None)
    if pending_id:
        req = PENDING_CLARIFICATION.pop(pending_id)
        req.description = f"{req.description} [{body.strip()}]"
        decision = engine.evaluate(req)
        if decision.action == "AUTO_APPROVE":
            send_auto_approval_notice(submitter, req.vendor, req.amount, decision.request_id)
            sync_approved_to_history()
            return build_twiml_reply(f"✅ Auto-approved.\nRef: {decision.request_id}")
        elif decision.action == "ESCALATE":
            send_approval_request(decision.escalation_message)
            return build_twiml_reply(f"📤 Sent to Sudhir.\nRef: {decision.request_id}")

    parts = [p.strip() for p in body.split(",")]
    if len(parts) < 5:
        return build_twiml_reply(
            "Format: Vendor, Amount, Category, Description, cash/upi\n"
            "Example: Swiggy, 850, dining, Dinner order, upi"
        )

    try:
        vendor = parts[0]
        amount = float(parts[1].replace("Rs", "").replace("rs", "").replace(",", "").strip())
        category = parts[2].lower().strip()
        description = parts[3]
        payment = parts[4].lower().strip()
        is_post_facto = len(parts) > 5 and "post" in parts[5].lower()
    except (ValueError, IndexError):
        return build_twiml_reply("Couldn't parse. Check format and try again.")

    req = ExpenseRequest(submitter=submitter, vendor=vendor, amount=amount,
                         category=category, description=description,
                         payment_method=payment, is_post_facto=is_post_facto)
    decision = engine.evaluate(req)

    if decision.action == "AUTO_APPROVE":
        send_auto_approval_notice(submitter, vendor, amount, decision.request_id)
        sync_approved_to_history()
        reply = f"✅ Auto-approved!\nRef: {decision.request_id}"
        if decision.budget_alert:
            reply += "\n⚠️ Category nearing monthly budget."
        return build_twiml_reply(reply)

    elif decision.action == "ESCALATE":
        send_approval_request(decision.escalation_message)
        reply = f"📤 Sent to Sudhir for approval.\nRef: {decision.request_id}"
        return build_twiml_reply(reply)

    elif decision.action == "PENDING_CLARIFICATION":
        PENDING_CLARIFICATION[decision.request_id] = req
        send_clarification_request(submitter, decision.follow_up_question,
                                   decision.follow_up_options, decision.request_id)
        return build_twiml_reply(f"❓ One question sent to clarify.\nRef: {decision.request_id}")

    return build_twiml_reply("Something went wrong. Please try again.")


@app.route("/webhook", methods=["POST"])
def webhook():
    incoming = parse_incoming(request.form)
    sender = incoming["from"]
    body = incoming["body"]
    if not body:
        return build_twiml_reply(""), 200
    if sender == SUDHIR:
        return app.response_class(_handle_sudhir_response(body), mimetype="text/xml")
    if sender in NUMBER_TO_NAME:
        return app.response_class(_handle_member_message(sender, body), mimetype="text/xml")
    return build_twiml_reply(""), 200


# ── SCHEDULER ────────────────────────────────────────────────────────────────

def _scheduled_reconciliation():
    try:
        run_reconciliation(notify_sudhir=True)
    except Exception as e:
        print(f"Reconciliation error: {e}")


if __name__ == "__main__":
    scheduler = BackgroundScheduler()
    scheduler.add_job(_scheduled_reconciliation, "interval", minutes=15)
    scheduler.start()
    port = int(os.getenv("PORT", 5000))
    print(f"Starting expense app on port {port}...")
    app.run(host="0.0.0.0", port=port, debug=False)
