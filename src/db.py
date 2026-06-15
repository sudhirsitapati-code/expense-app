"""
db.py — PostgreSQL-backed key-value store for persistent JSON data.
Falls back to local file storage when DATABASE_URL is not set (local dev).
"""

import json
import os

DATABASE_URL = os.getenv("DATABASE_URL", "")

# ── PostgreSQL path ───────────────────────────────────────────────────────────

def _conn():
    import psycopg2
    return psycopg2.connect(DATABASE_URL)


def init_db():
    """Create kv_store table if it doesn't exist. Call once at startup."""
    if not DATABASE_URL:
        return
    try:
        with _conn() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS kv_store (
                        key        TEXT PRIMARY KEY,
                        value      JSONB NOT NULL DEFAULT '[]'::jsonb,
                        updated_at TIMESTAMP DEFAULT NOW()
                    )
                """)
            conn.commit()
    except Exception as e:
        print(f"DB init error: {e}")


def load(key: str, default=None):
    """Load a JSON value by key. Returns default if not found."""
    if default is None:
        default = []
    if not DATABASE_URL:
        return _file_load(key, default)
    try:
        with _conn() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT value FROM kv_store WHERE key = %s", (key,))
                row = cur.fetchone()
                return row[0] if row else default
    except Exception as e:
        print(f"DB load error ({key}): {e}")
        return _file_load(key, default)


def save(key: str, value):
    """Persist a JSON-serialisable value under key."""
    if not DATABASE_URL:
        return _file_save(key, value)
    try:
        with _conn() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    INSERT INTO kv_store (key, value, updated_at)
                    VALUES (%s, %s::jsonb, NOW())
                    ON CONFLICT (key) DO UPDATE
                        SET value = EXCLUDED.value, updated_at = NOW()
                """, (key, json.dumps(value)))
            conn.commit()
    except Exception as e:
        print(f"DB save error ({key}): {e}")
        _file_save(key, value)


# ── File fallback (local dev) ─────────────────────────────────────────────────

_BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_DATA_DIR = os.path.join(_BASE_DIR, "data")

_KEY_TO_FILE = {
    "approval_log":            "approval_log.json",
    "reconcile_log":           "reconcile_log.json",
    "master_ledger":           "master_ledger.json",
    "icici_transactions":      "icici_transactions.json",
    "processed_gmail_ids":     "processed_gmail_ids.json",
    "processed_statement_ids": "processed_statement_ids.json",
}


def _file_path(key: str) -> str:
    filename = _KEY_TO_FILE.get(key, f"{key}.json")
    return os.path.join(_DATA_DIR, filename)


def _file_load(key: str, default):
    path = _file_path(key)
    if not os.path.exists(path):
        return default
    try:
        with open(path) as f:
            return json.load(f)
    except Exception:
        return default


def _file_save(key: str, value):
    os.makedirs(_DATA_DIR, exist_ok=True)
    path = _file_path(key)
    with open(path, "w") as f:
        json.dump(value, f, indent=2)
