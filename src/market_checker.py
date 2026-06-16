"""
market_checker.py
Four-layer market check using Azure OpenAI.
"""

import json
import os
from typing import Optional
from openai import AzureOpenAI

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
DATA_DIR = os.path.join(BASE_DIR, "data")

with open(os.path.join(CONFIG_DIR, "price_book.json")) as f:
    PRICE_BOOK = json.load(f)["services"]

ACC26_HISTORY_PATH = os.path.join(DATA_DIR, "acc26_history.json")


class MarketChecker:
    def __init__(self):
        self.client = AzureOpenAI(
            api_key=os.getenv("AZURE_OPENAI_KEY"),
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
            api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-12-01-preview")
        )
        self.deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-5.5")

    def check(self, description: str, vendor: str, amount: float, category: str) -> dict:
        result = self._check_price_book(description, vendor, amount)
        if result:
            return result
        result = self._check_acc26_history(description, vendor, amount, category)
        if result:
            return result
        result = self._ai_estimate(description, vendor, amount, category)
        if result:
            return result
        return {"status": "unknown", "rate_range": None, "source": "no_data"}

    def _check_price_book(self, description: str, vendor: str, amount: float) -> Optional[dict]:
        text = (description + " " + vendor).lower()
        for service_key, service in PRICE_BOOK.items():
            for kw in service.get("keywords", []):
                if kw.lower() in text:
                    min_r, max_r = service["min"], service["max"]
                    rate_range = f"Rs {min_r:,}–{max_r:,} {service['unit']}"
                    if amount > max_r * 1.3:
                        return {"status": "very_high", "rate_range": rate_range, "source": "price_book"}
                    elif amount > max_r:
                        return {"status": "high", "rate_range": rate_range, "source": "price_book"}
                    else:
                        return {"status": "ok", "rate_range": rate_range, "source": "price_book"}
        return None

    def _check_acc26_history(self, description: str, vendor: str, amount: float, category: str) -> Optional[dict]:
        if not os.path.exists(ACC26_HISTORY_PATH):
            return None
        try:
            with open(ACC26_HISTORY_PATH) as f:
                history = json.load(f)
        except Exception:
            return None
        text = (description + " " + vendor).lower()
        matches = []
        for entry in history:
            hist_desc = (entry.get("description", "") + " " + entry.get("vendor", "")).lower()
            common_words = set(text.split()) & set(hist_desc.split())
            if len(common_words) >= 2 and entry.get("category") == category:
                matches.append(entry["amount"])
        if not matches:
            return None
        avg = sum(matches) / len(matches)
        rate_range = f"Rs {min(matches):,.0f}–{max(matches):,.0f} (historical)"
        if amount > avg * 1.3:
            return {"status": "very_high", "rate_range": rate_range, "source": "acc26_history"}
        elif amount > avg * 1.1:
            return {"status": "high", "rate_range": rate_range, "source": "acc26_history"}
        else:
            return {"status": "ok", "rate_range": rate_range, "source": "acc26_history"}

    def _ai_estimate(self, description: str, vendor: str, amount: float, category: str) -> Optional[dict]:
        prompt = f"""You are a Mumbai household expense advisor.
Assess whether this expense is reasonably priced for Mumbai in 2025-2026:
- Service/Item: {description}
- Vendor: {vendor}
- Category: {category}
- Amount charged: Rs {amount:,.0f}

Reply with JSON only:
{{"status": "ok"|"high"|"very_high"|"unknown", "rate_range": "Rs X–Y per [unit] (estimated)", "reasoning": "one line"}}

"ok" = within typical Mumbai market rate
"high" = 10-30% above market
"very_high" = >30% above market
"unknown" = not enough info

Reply ONLY with JSON."""

        try:
            response = self.client.chat.completions.create(
                model=self.deployment,
                max_completion_tokens=200,
                messages=[
                    {"role": "system", "content": "You are a helpful assistant that replies only with JSON."},
                    {"role": "user", "content": prompt}
                ]
            )
            text = response.choices[0].message.content.strip()
            # Strip markdown code blocks if present
            if text.startswith("```"):
                text = text.split("```")[1]
                if text.startswith("json"):
                    text = text[4:]
            result = json.loads(text)
            result["source"] = "ai_estimate"
            return result
        except Exception:
            return None
