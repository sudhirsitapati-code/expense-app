"""
gmail_utils.py
Shared helper to get Gmail credentials — works both locally (from file)
and on Railway (from GMAIL_TOKEN_JSON env var).
"""

import json
import os
import tempfile

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def get_credentials() -> Credentials:
    """Return valid Gmail credentials from file or env var."""
    token_json = os.getenv("GMAIL_TOKEN_JSON")
    token_file = os.getenv("GMAIL_TOKEN_FILE", "config/gmail_token.json")

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    token_path = os.path.join(base_dir, token_file)

    if token_json:
        # Railway: token is in env var — write to a temp file
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
        tmp.write(token_json)
        tmp.close()
        creds = Credentials.from_authorized_user_file(tmp.name, SCOPES)
        os.unlink(tmp.name)
    elif os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    else:
        raise FileNotFoundError(
            f"Gmail token not found. Set GMAIL_TOKEN_JSON env var or place token at {token_path}"
        )

    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        # Save refreshed token back
        if not token_json and os.path.exists(token_path):
            with open(token_path, "w") as f:
                f.write(creds.to_json())

    return creds
