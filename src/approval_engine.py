"""
approval_engine.py
Core classification + routing logic — uses Azure OpenAI.
"""

import json
import os
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional
from openai import AzureOpenAI
from src.market_checker import MarketChecker

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
DATA_DIR = os.path.join(BASE_DIR, "data")

with open(os.path.join(CONFIG_DIR, "approved_recurring.json")) as f:
    APPROVED_RECURRING = json.load(f)["recurring"]

with open(os.path.join(CONFIG_DIR, "budget_fy27.json")) as f:
    BUDGET = json.load(f)["monthly"]

APPROVAL_LOG_PATH = os.path.join(DATA_DIR, "approval_log.json")

from src import db as _db


@dataclass
class ExpenseRequest:
    submitter: str
    vendor: str
    amount: float
    category: str
    description: str
    payment_method: str
    is_post_facto: bool = False
    request_id: str = ""
    timestamp: str = ""

    def __post_init__(self):
        if not self.request_id:
            self.request_id = f"REQ-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
        if not self.timestamp:
            self.timestamp = datetime.now().isoformat()


@dataclass
class ApprovalDecision:
    request_id: str
    action: str
    reason: str
    market_status: Optional[str] = None
    market_rate: Optional[str] = None
    budget_alert: bool = False
    escalation_message: Optional[str] = None
    follow_up_question: Optional[str] = None
    follow_up_options: list = field(default_factory=list)
    confirmed_paid: bool = False  # True for post-facto items already paid


class ApprovalEngine:
    def __init__(self):
        self.client = AzureOpenAI(
            api_key=os.getenv("AZURE_OPENAI_KEY"),
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
            api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-12-01-preview")
        )
        self.deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-5.5")
        self.market_checker = MarketChecker()

    def _match_recurring(self, req: ExpenseRequest) -> Optional[dict]:
        vendor_lower = req.vendor.lower()
        for rec in APPROVED_RECURRING:
            for kw in rec["payee_keywords"]:
                if kw.lower() in vendor_lower:
                    if rec["amount_min"] <= req.amount <= rec["amount_max"]:
                        return rec
        return None

    def _budget_alert(self, category: str, amount: float) -> bool:
        monthly = BUDGET.get(category, 0)
        if monthly == 0:
            return False
        current = self._get_current_month_spend(category)
        return (current + amount) > monthly * 0.80

    def _get_current_month_spend(self, category: str) -> float:
        try:
            log = _db.load("approval_log")
            month_prefix = datetime.now().strftime("%Y-%m")
            return sum(
                e.get("amount", 0) for e in log
                if e.get("category") == category
                and e.get("timestamp", "").startswith(month_prefix)
                and e.get("action") in ("AUTO_APPROVE", "APPROVED")
            )
        except Exception:
            return 0.0

    def _needs_clarification(self, req: ExpenseRequest) -> Optional[dict]:
        if req.amount < 5000:
            return None
        prompt = f"""You are a household expense approval assistant.
Vincent submitted this expense:
- Vendor: {req.vendor}
- Amount: Rs {req.amount:,.0f}
- Category: {req.category}
- Description: {req.description}

Is the description specific enough to judge whether the price is reasonable?
If YES: reply with JSON {{"clear": true}}
If NO: reply with JSON {{"clear": false, "question": "<one short question>", "options": ["<opt1>", "<opt2>", "<opt3>"]}}
Options must be short chip-style answers (max 4 words each).
Reply ONLY with JSON."""

        try:
            response = self.client.chat.completions.create(
                model=self.deployment,
                max_completion_tokens=200,
                messages=[
                    {"role": "system", "content": "Reply only with JSON."},
                    {"role": "user", "content": prompt}
                ]
            )
            text = response.choices[0].message.content.strip()
            if text.startswith("```"):
                text = text.split("```")[1]
                if text.startswith("json"):
                    text = text[4:]
            result = json.loads(text)
            if result.get("clear"):
                return None
            return {"question": result["question"], "options": result.get("options", [])}
        except Exception:
            return None

    def evaluate(self, req: ExpenseRequest) -> ApprovalDecision:
        decision = ApprovalDecision(request_id=req.request_id, action="", reason="")

        # Rule 1: Recurring — always auto-approve
        matched = self._match_recurring(req)
        if matched:
            decision.action = "AUTO_APPROVE"
            decision.reason = f"Matched recurring: {matched['description']}"
            decision.budget_alert = self._budget_alert(req.category, req.amount)
            self._save_to_log(req, decision)
            return decision

        # Rule 2: Post-facto (already paid) — log immediately as confirmed expense.
        # No point asking for approval on money already spent.
        if req.is_post_facto:
            decision.action = "AUTO_APPROVE"
            decision.reason = "Post-facto — already paid, logged as confirmed expense"
            decision.budget_alert = self._budget_alert(req.category, req.amount)
            decision.confirmed_paid = True
            self._save_to_log(req, decision)
            return decision

        # Rule 3: Small amount — auto-approve
        if req.amount < 5000 and not self._budget_alert(req.category, req.amount):
            decision.action = "AUTO_APPROVE"
            decision.reason = "Amount < Rs 5,000 and within budget"
            self._save_to_log(req, decision)
            return decision

        # Rule 4: Market check
        market_result = self.market_checker.check(
            description=req.description, vendor=req.vendor,
            amount=req.amount, category=req.category
        )
        decision.market_status = market_result.get("status", "unknown")
        decision.market_rate = market_result.get("rate_range")
        decision.budget_alert = self._budget_alert(req.category, req.amount)

        # Rule 5: Thresholds
        if req.amount > 30000:
            decision.action = "ESCALATE"
            decision.reason = "Amount > Rs 30,000"
        elif decision.market_status == "very_high":
            decision.action = "ESCALATE"
            decision.reason = f"Quote >30% above market rate ({decision.market_rate})"
        elif req.amount >= 5000:
            decision.action = "ESCALATE"
            decision.reason = "Amount Rs 5,000–30,000: requires approval"
        else:
            decision.action = "AUTO_APPROVE"
            decision.reason = "Within threshold, market rate OK"

        if decision.action == "ESCALATE":
            market_note = ""
            if decision.market_status == "very_high":
                market_note = f"\n⚠️ Quote is HIGH vs market ({decision.market_rate})"
            elif decision.market_status == "high":
                market_note = f"\n⚠️ Slightly above market ({decision.market_rate})"
            elif decision.market_rate:
                market_note = f"\nMarket rate: {decision.market_rate}"
            budget_note = "\n📊 Category near monthly budget limit" if decision.budget_alert else ""
            decision.escalation_message = self._build_escalation_message(req, market_note + budget_note)

        self._save_to_log(req, decision)
        return decision

    def _build_escalation_message(self, req: ExpenseRequest, extra_note: str = "") -> str:
        payment_tag = "💵 Cash" if req.payment_method == "cash" else "🏦 SBI-3152"
        post_tag = " [POST-FACTO]" if req.is_post_facto else ""
        return (
            f"💰 *Expense Approval Required{post_tag}*\n"
            f"From: {req.submitter.title()}\n"
            f"Vendor: {req.vendor}\n"
            f"Amount: Rs {req.amount:,.0f}\n"
            f"Category: {req.category}\n"
            f"Purpose: {req.description}\n"
            f"Payment: {payment_tag}{extra_note}\n\n"
            f"Reply *Y* to approve | *N* to reject | *L [amount]* for lower\n"
            f"Ref: {req.request_id}"
        )

    def _save_to_log(self, req: ExpenseRequest, decision: ApprovalDecision):
        log = _db.load("approval_log")
        entry = {
            "request_id": req.request_id,
            "timestamp": req.timestamp,
            "submitter": req.submitter,
            "vendor": req.vendor,
            "amount": req.amount,
            "category": req.category,
            "description": req.description,
            "payment_method": req.payment_method,
            "is_post_facto": req.is_post_facto,
            "action": decision.action,
            "reason": decision.reason,
            "market_status": decision.market_status,
            "market_rate": decision.market_rate,
            "budget_alert": decision.budget_alert,
            "confirmed_paid": decision.confirmed_paid,
            "confirmed_at": datetime.now().isoformat() if decision.confirmed_paid else None,
            "confirmed_by": "post_facto" if decision.confirmed_paid else None,
        }
        log.append(entry)
        _db.save("approval_log", log)

    def update_log_with_sudhir_response(self, request_id: str, response: str):
        log = _db.load("approval_log")
        for entry in log:
            if entry["request_id"] == request_id:
                resp_upper = response.strip().upper()
                if resp_upper == "Y":
                    entry["action"] = "APPROVED"
                    entry["sudhir_response"] = "Y"
                elif resp_upper == "N":
                    entry["action"] = "REJECTED"
                    entry["sudhir_response"] = "N"
                elif resp_upper.startswith("L"):
                    try:
                        lower_amount = float(resp_upper[1:].strip())
                        entry["action"] = "APPROVED_LOWER"
                        entry["approved_amount"] = lower_amount
                        entry["sudhir_response"] = f"L {lower_amount}"
                    except ValueError:
                        pass
                entry["response_timestamp"] = datetime.now().isoformat()
                break
        _db.save("approval_log", log)
