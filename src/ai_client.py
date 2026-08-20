"""
ai_client.py
Centralized AI client factory. Uses the personal OpenAI key as the primary
account, falling back to the Godrej-issued Azure OpenAI key if it's not set.
Recovering from a lost or revoked key is then an .env change, not a hunt
through every call site.
"""

import os
from openai import AzureOpenAI, OpenAI

DEFAULT_AZURE_DEPLOYMENT = "gpt-5.5"
DEFAULT_OPENAI_MODEL = "gpt-4o"


def get_ai_client():
    """Returns (client, model). Prefers standard OpenAI, falls back to Azure OpenAI."""
    if os.getenv("OPENAI_API_KEY"):
        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        model = os.getenv("OPENAI_MODEL", DEFAULT_OPENAI_MODEL)
        return client, model

    if os.getenv("AZURE_OPENAI_KEY"):
        client = AzureOpenAI(
            api_key=os.getenv("AZURE_OPENAI_KEY"),
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
            api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-12-01-preview"),
        )
        model = os.getenv("AZURE_OPENAI_DEPLOYMENT", DEFAULT_AZURE_DEPLOYMENT)
        return client, model

    raise RuntimeError("No AI credentials found — set OPENAI_API_KEY or AZURE_OPENAI_KEY in .env")
