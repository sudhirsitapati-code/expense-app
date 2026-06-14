"""
whatsapp_handler.py
Sends and receives WhatsApp messages via Twilio.
"""

import os
from twilio.rest import Client
from twilio.twiml.messaging_response import MessagingResponse

ACCOUNT_SID = os.getenv("TWILIO_ACCOUNT_SID")
AUTH_TOKEN = os.getenv("TWILIO_AUTH_TOKEN")
FROM_NUMBER = os.getenv("TWILIO_WHATSAPP_FROM", "whatsapp:+14155238886")

SUDHIR = os.getenv("SUDHIR_WHATSAPP")
VINCENT = os.getenv("VINCENT_WHATSAPP")
KETKI = os.getenv("KETKI_WHATSAPP")
SANTOSH = os.getenv("SANTOSH_WHATSAPP")

HOUSEHOLD_MEMBERS = {
    "vincent": VINCENT,
    "ketki": KETKI,
    "santosh": SANTOSH,
}


def _client() -> Client:
    return Client(ACCOUNT_SID, AUTH_TOKEN)


def send_message(to: str, body: str) -> str:
    """Send a WhatsApp message. Returns message SID."""
    msg = _client().messages.create(from_=FROM_NUMBER, to=to, body=body)
    return msg.sid


def send_to_sudhir(body: str) -> str:
    return send_message(SUDHIR, body)


def send_approval_request(escalation_message: str) -> str:
    """Send an expense approval request to Sudhir."""
    return send_to_sudhir(escalation_message)


def send_auto_approval_notice(submitter: str, vendor: str, amount: float, request_id: str):
    """Notify submitter that their expense was auto-approved."""
    submitter_number = HOUSEHOLD_MEMBERS.get(submitter.lower())
    if not submitter_number:
        return
    msg = (
        f"✅ Your expense has been auto-approved.\n"
        f"Vendor: {vendor}\n"
        f"Amount: Rs {amount:,.0f}\n"
        f"Ref: {request_id}"
    )
    send_message(submitter_number, msg)


def send_approval_result(submitter: str, vendor: str, amount: float,
                         approved: bool, request_id: str, approved_amount: float = None):
    """Notify submitter of Sudhir's approval/rejection decision."""
    submitter_number = HOUSEHOLD_MEMBERS.get(submitter.lower())
    if not submitter_number:
        return
    if approved:
        final_amount = approved_amount or amount
        msg = (
            f"✅ Expense APPROVED by Sudhir.\n"
            f"Vendor: {vendor}\n"
            f"Amount: Rs {final_amount:,.0f}\n"
            f"Ref: {request_id}"
        )
    else:
        msg = (
            f"❌ Expense REJECTED by Sudhir.\n"
            f"Vendor: {vendor}\n"
            f"Amount: Rs {amount:,.0f}\n"
            f"Ref: {request_id}"
        )
    send_message(submitter_number, msg)


def send_clarification_request(submitter: str, question: str,
                                options: list, request_id: str):
    """Ask submitter a clarifying question before processing."""
    submitter_number = HOUSEHOLD_MEMBERS.get(submitter.lower())
    if not submitter_number:
        return
    options_text = "\n".join(f"{i+1}. {opt}" for i, opt in enumerate(options))
    msg = (
        f"❓ {question}\n\n"
        f"{options_text}\n\n"
        f"Reply with the number or type your answer.\n"
        f"Ref: {request_id}"
    )
    send_message(submitter_number, msg)


def build_twiml_reply(message: str) -> str:
    """Build a TwiML response string for Twilio webhook."""
    resp = MessagingResponse()
    resp.message(message)
    return str(resp)


def parse_incoming(form_data: dict) -> dict:
    """Extract relevant fields from a Twilio incoming webhook POST."""
    return {
        "from": form_data.get("From", ""),
        "body": form_data.get("Body", "").strip(),
        "message_sid": form_data.get("MessageSid", ""),
    }
