"""
app.py
Flask app — WhatsApp webhook for expense submission and approval responses.
"""

import json
import os

from apscheduler.schedulers.background import BackgroundScheduler
from dotenv import load_dotenv
from flask import Flask, request

from src.approval_engine import ApprovalEngine, ExpenseRequest
from src.reconcile import run_reconciliation
from src.acc27_writer import sync_approved_to_history, export_monthly_excel
from src.whatsapp_handler import (
    build_twiml_reply,
    parse_incoming,
    send_approval_request,
    send_approval_result,
    send_auto_approval_notice,
    send_clarification_request,
    SUDHIR,
    HOUSEHOLD_MEMBERS,
)

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-me")

engine = ApprovalEngine()

# In-memory store for pending clarifications: request_id -> ExpenseRequest
PENDING_CLARIFICATION: dict[str, ExpenseRequest] = {}

# Reverse lookup: whatsapp_number -> name
NUMBER_TO_NAME = {v: k for k, v in HOUSEHOLD_MEMBERS.items() if v}


def _submitter_from_number(number: str) -> str:
    return NUMBER_TO_NAME.get(number, "unknown")


def _handle_sudhir_response(body: str):
    """Process Y / N / L <amount> replies from Sudhir."""
    parts = body.strip().split()
    command = parts[0].upper()

    # Load approval log to find the most recent ESCALATE entry awaiting response
    data_dir = os.path.join(os.path.dirname(__file__), "data")
    log_path = os.path.join(data_dir, "approval_log.json")
    if not os.path.exists(log_path):
        return build_twiml_reply("No pending approvals found.")

    with open(log_path) as f:
        log = json.load(f)

    pending = [
        e for e in log
        if e.get("action") == "ESCALATE" and "sudhir_response" not in e
    ]
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
        return build_twiml_reply(f"✅ Approved. {submitter.title()} has been notified.\nRef: {request_id}")

    elif command == "N":
        send_approval_result(submitter, vendor, amount, approved=False, request_id=request_id)
        return build_twiml_reply(f"❌ Rejected. {submitter.title()} has been notified.\nRef: {request_id}")

    elif command == "L" and len(parts) > 1:
        try:
            lower_amount = float(parts[1].replace(",", ""))
            send_approval_result(submitter, vendor, amount, approved=True,
                                 request_id=request_id, approved_amount=lower_amount)
            sync_approved_to_history()
            return build_twiml_reply(
                f"✅ Approved at Rs {lower_amount:,.0f}. {submitter.title()} has been notified.\nRef: {request_id}"
            )
        except ValueError:
            return build_twiml_reply("Invalid amount. Use: L 5000")

    return build_twiml_reply("Use Y / N / L <amount> to respond.")


def _handle_member_message(sender: str, body: str):
    """Parse an expense submission from a household member."""
    submitter = _submitter_from_number(sender)

    # Check if this is a clarification reply
    pending_id = None
    for req_id, req in PENDING_CLARIFICATION.items():
        if req.submitter.lower() == submitter.lower():
            pending_id = req_id
            break

    if pending_id:
        req = PENDING_CLARIFICATION.pop(pending_id)
        req.description = f"{req.description} [{body.strip()}]"
        decision = engine.evaluate(req)

        if decision.action == "AUTO_APPROVE":
            send_auto_approval_notice(submitter, req.vendor, req.amount, decision.request_id)
            sync_approved_to_history()
            return build_twiml_reply(f"✅ Auto-approved after clarification.\nRef: {decision.request_id}")
        elif decision.action == "ESCALATE":
            send_approval_request(decision.escalation_message)
            return build_twiml_reply(f"📤 Sent to Sudhir for approval.\nRef: {decision.request_id}")

    # Parse new expense: expected format:
    # <vendor>, <amount>, <category>, <description>, <cash|upi>, [post-facto]
    parts = [p.strip() for p in body.split(",")]
    if len(parts) < 5:
        return build_twiml_reply(
            "Please send expense in this format:\n"
            "Vendor, Amount, Category, Description, cash/upi\n\n"
            "Example:\nSwiggy, 850, dining, Dinner order, upi"
        )

    try:
        vendor = parts[0]
        amount = float(parts[1].replace("rs", "").replace("Rs", "").replace(",", "").strip())
        category = parts[2].lower().strip()
        description = parts[3]
        payment = parts[4].lower().strip()
        is_post_facto = len(parts) > 5 and "post" in parts[5].lower()
    except (ValueError, IndexError):
        return build_twiml_reply("Couldn't parse your expense. Please check the format and try again.")

    req = ExpenseRequest(
        submitter=submitter,
        vendor=vendor,
        amount=amount,
        category=category,
        description=description,
        payment_method=payment,
        is_post_facto=is_post_facto,
    )

    decision = engine.evaluate(req)

    if decision.action == "AUTO_APPROVE":
        send_auto_approval_notice(submitter, vendor, amount, decision.request_id)
        sync_approved_to_history()
        reply = f"✅ Auto-approved!\nRef: {decision.request_id}"
        if decision.budget_alert:
            reply += "\n⚠️ Note: Category nearing monthly budget limit."
        return build_twiml_reply(reply)

    elif decision.action == "ESCALATE":
        send_approval_request(decision.escalation_message)
        reply = f"📤 Sent to Sudhir for approval.\nRef: {decision.request_id}"
        if decision.budget_alert:
            reply += "\n⚠️ Note: Category nearing monthly budget limit."
        return build_twiml_reply(reply)

    elif decision.action == "PENDING_CLARIFICATION":
        PENDING_CLARIFICATION[decision.request_id] = req
        send_clarification_request(
            submitter, decision.follow_up_question,
            decision.follow_up_options, decision.request_id
        )
        return build_twiml_reply(f"❓ One quick question sent to clarify your expense.\nRef: {decision.request_id}")

    return build_twiml_reply("Something went wrong. Please try again.")


@app.route("/webhook", methods=["POST"])
def webhook():
    incoming = parse_incoming(request.form)
    sender = incoming["from"]
    body = incoming["body"]

    if not body:
        return build_twiml_reply(""), 200

    if sender == SUDHIR:
        return app.response_class(
            _handle_sudhir_response(body), mimetype="text/xml"
        )

    if sender in NUMBER_TO_NAME:
        return app.response_class(
            _handle_member_message(sender, body), mimetype="text/xml"
        )

    return build_twiml_reply(""), 200


@app.route("/health", methods=["GET"])
def health():
    return {"status": "ok"}, 200


@app.route("/export", methods=["GET"])
def export():
    path = export_monthly_excel()
    return {"status": "exported", "file": path}, 200


def _scheduled_reconciliation():
    try:
        unmatched = run_reconciliation(notify_sudhir=True)
        if unmatched:
            print(f"Reconciliation: {len(unmatched)} unmatched debit(s) flagged.")
    except Exception as e:
        print(f"Reconciliation error: {e}")


if __name__ == "__main__":
    scheduler = BackgroundScheduler()
    scheduler.add_job(_scheduled_reconciliation, "interval", hours=6)
    scheduler.start()

    port = int(os.getenv("PORT", 5000))
    print(f"Starting expense app on port {port}...")
    app.run(host="0.0.0.0", port=port, debug=False)
