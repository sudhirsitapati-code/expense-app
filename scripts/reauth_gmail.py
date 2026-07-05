"""
Re-authorize Gmail OAuth and write a fresh gmail_token.json.
Run this locally whenever you get invalid_grant errors.

Usage:
  cd /Users/sudhirsitapati/Desktop/expense-app
  python scripts/reauth_gmail.py
"""

import json
import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CREDS_FILE = os.path.join(BASE_DIR, "config", "gmail_credentials.json")
TOKEN_FILE  = os.path.join(BASE_DIR, "config", "gmail_token.json")

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

try:
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.oauth2.credentials import Credentials
except ImportError:
    print("Missing packages. Run:  pip install google-auth-oauthlib google-auth-httplib2")
    sys.exit(1)

if not os.path.exists(CREDS_FILE):
    print(f"ERROR: credentials not found at {CREDS_FILE}")
    sys.exit(1)

# Delete stale token first
if os.path.exists(TOKEN_FILE):
    os.remove(TOKEN_FILE)
    print(f"Removed stale token: {TOKEN_FILE}")

flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
creds = flow.run_local_server(port=0)

with open(TOKEN_FILE, "w") as f:
    f.write(creds.to_json())

print(f"\nDone! Fresh token written to: {TOKEN_FILE}")
print("\nNow copy it into Railway as the GMAIL_TOKEN_JSON env var:")
print("=" * 60)
print(open(TOKEN_FILE).read())
