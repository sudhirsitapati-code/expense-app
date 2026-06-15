"""
app.py
Flask app — all routes: WhatsApp webhook, HTML screens, JSON APIs.
"""

import json
import os
from datetime import datetime

from apscheduler.schedulers.background import BackgroundScheduler
from dotenv import load_dotenv
from flask import Flask, request, render_template, jsonify, session, redirect, url_for

from src.approval_engine import ApprovalEngine, ExpenseRequest
from src.reconcile import run_reconciliation
from src.acc27_writer import sync_approved_to_history, export_monthly_excel
from src.icici_statement_parser import fetch_and_parse_statements
from src.whatsapp_handler import (
    build_twiml_reply, parse_incoming,
    send_approval_request, send_approval_result,
    send_auto_approval_notice, send_clarification_request,
    SUDHIR, HOUSEHOLD_MEMBERS,
)

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-me")

engine = ApprovalEngine()

BASE_DIR = os.path.dirname(__file__)
DATA_DIR = os.path.join(BASE_DIR, "data")
CONFIG_DIR = os.path.join(BASE_DIR, "config")

APPROVAL_LOG = os.path.join(DATA_DIR, "approval_log.json")
RECONCILE_LOG = os.path.join(DATA_DIR, "reconcile_log.json")
TRANSACTIONS_PATH = os.path.join(DATA_DIR, "icici_transactions.json")

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

def _load_json(path):
    if not os.path.exists(path):
        return []
    with open(path) as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def _save_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(data, f, indent=2)


def _update_log_entry(request_id: str, updates: dict):
    log = _load_json(APPROVAL_LOG)
    for entry in log:
        if entry.get("request_id") == request_id:
            entry.update(updates)
            break
    _save_json(APPROVAL_LOG, log)


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
        budget_monthly = json.load(f)["monthly"]

    # FY26 actuals by ACC26 heading — sourced from ExpenseSummary sheet of ACC26ver5_MASTER.xlsx
    # Stored here so they survive Railway's ephemeral filesystem (data/ is gitignored)
    FY26_ACTUALS = {
        "Misc": 189000, "Cash": 258000, "Electricity & Gas": 393000,
        "Groceries": 1689000, "Staff Salary": 2631000,
        "Alcohol": 75000, "Wellness": 388000,
        "Clothes": 391000, "Gifts": 100000, "Medical": 591000,
        "Amma": 114000, "Ketki": 1545000, "Children Education": 3033000,
        "Charity": 1526000, "Uspaar": 1304000,
        "Holiday": 2189000, "Eating Out": 520000, "Entertainment": 226000,
        "Malhar": 965000, "Maintenance Expense": 88000, "Home office": 234000,
        "One Time Charge": 771000, "Kalpataru Maintenance": 549000,
        "Financial Expense / OD Interest": 9000, "Insurance": 364000,
        "Home Loan": 13067000, "Tax": 45100000,
    }
    # Merge with any locally saved file (allows future updates without redeploying)
    fy26_by_heading = dict(FY26_ACTUALS)
    local_fy26 = _load_json(os.path.join(DATA_DIR, "fy26_actuals.json"))
    if isinstance(local_fy26, dict):
        fy26_by_heading.update(local_fy26)  # handle empty/old format

    log = _load_json(APPROVAL_LOG)

    # ── Period months ─────────────────────────────────────────────────────────
    now = datetime.now()
    cal_month = now.month
    fy_start_year = now.year if cal_month >= 4 else now.year - 1

    def _period_months(period):
        if period == "month":
            return [now.strftime("%Y-%m")]
        elif period == "quarter":
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

    # ── Mappings ──────────────────────────────────────────────────────────────
    # App category → ACC26 headings (for pulling FY26 data)
    CAT_TO_HEADINGS = {
        "groceries":     ["Groceries"],
        "staff":         ["Staff Salary"],
        "utilities":     ["Electricity & Gas"],
        "miscellaneous": ["Misc"],
        "personal_care": ["Wellness"],
        "clothing":      ["Clothes"],
        "gifts":         ["Gifts"],
        "medical":       ["Medical"],
        "education":     ["Children Education"],
        "dining":        ["Eating Out"],
        "entertainment": ["Entertainment"],
        "transport":     ["Holiday"],
        "maintenance":   ["Maintenance Expense", "Kalpataru Maintenance"],
        "home_repair":   ["One Time Charge", "Home office"],
    }
    # App category → super-category
    SUPER_CAT = {
        "groceries": "Household", "staff": "Household",
        "utilities": "Household", "miscellaneous": "Household",
        "personal_care": "Personal",
        "clothing": "Family", "gifts": "Family",
        "medical": "Family", "education": "Family",
        "dining": "Lifestyle", "entertainment": "Lifestyle", "transport": "Lifestyle",
        "maintenance": "Property", "home_repair": "Property",
    }
    SUPER_ORDER = ["Household", "Personal", "Family", "Giving", "Lifestyle", "Property", "Financial"]

    # ── FY26 full-year per app-category (sum of mapped headings) ─────────────
    def _fy26_for_cat(cat):
        return sum(fy26_by_heading.get(h, 0) for h in CAT_TO_HEADINGS.get(cat, []))

    # ── FY27 budget: full year for ytd, scaled for month/quarter ─────────────
    def _budget_for_cat(cat):
        monthly = budget_monthly.get(cat, 0)
        return monthly * 12 if period == "ytd" else monthly * n_months

    # ── FY27 actual for the period ────────────────────────────────────────────
    fy27_actual: dict = {}
    for e in log:
        if e.get("action") not in ("AUTO_APPROVE", "APPROVED", "APPROVED_LOWER"):
            continue
        if (e.get("timestamp") or "")[:7] not in period_months:
            continue
        cat = e.get("category", "miscellaneous")
        amt = e.get("approved_amount") or e.get("amount", 0)
        fy27_actual[cat] = fy27_actual.get(cat, 0) + amt

    # ── Build grouped rows ────────────────────────────────────────────────────
    by_super: dict = {s: [] for s in SUPER_ORDER}
    for cat in sorted(budget_monthly.keys()):
        super_cat = SUPER_CAT.get(cat, "Household")
        fy26 = round(_fy26_for_cat(cat))
        budget = _budget_for_cat(cat)
        actual = round(fy27_actual.get(cat, 0))
        pct = round(actual / budget * 100) if budget else 0
        by_super.setdefault(super_cat, []).append({
            "category": cat,
            "fy26_actual": fy26,
            "fy27_budget": budget,
            "fy27_actual": actual,
            "pct": pct,
        })

    # Add FY26-only super-categories (Giving, Financial) as reference rows
    HEADING_SUPER = {
        "Charity": "Giving", "Uspaar": "Giving",
        "Financial Expense / OD Interest": "Financial",
        "Insurance": "Financial", "Home Loan": "Financial", "Tax": "Financial",
        "Amma": "Family", "Ketki": "Family",
        "Malhar": "Property", "Cash": "Household", "Alcohol": "Personal",
    }
    already_mapped = {h for cat in CAT_TO_HEADINGS.values() for h in cat}
    for heading, fy26_amt in fy26_by_heading.items():
        if heading in already_mapped:
            continue
        super_cat = HEADING_SUPER.get(heading, "Household")
        by_super.setdefault(super_cat, []).append({
            "category": heading,
            "fy26_actual": round(fy26_amt),
            "fy27_budget": 0,
            "fy27_actual": 0,
            "pct": 0,
        })

    groups = []
    grand = {"fy26": 0, "budget": 0, "actual": 0}
    for super_cat in SUPER_ORDER:
        rows = by_super.get(super_cat, [])
        if not rows:
            continue
        sub = {
            "fy26":   sum(r["fy26_actual"] for r in rows),
            "budget": sum(r["fy27_budget"] for r in rows),
            "actual": sum(r["fy27_actual"] for r in rows),
        }
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


@app.route("/api/mark-paid", methods=["POST"])
def api_mark_paid():
    """Manually mark a cash expense as confirmed paid."""
    data = request.get_json()
    request_id = data.get("request_id")
    _update_log_entry(request_id, {"confirmed_paid": True, "confirmed_at": datetime.now().isoformat()})
    return jsonify({"status": "ok"})


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
