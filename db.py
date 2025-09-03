# -*- coding: utf-8 -*-
import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Dict, Any
from config import DB_PATH

DDL_MAIL_LOG = """
CREATE TABLE IF NOT EXISTS mail_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  email TEXT,
  company TEXT,
  token TEXT UNIQUE,
  subject TEXT,
  entryid TEXT,
  conversation_id TEXT,
  sent_on TEXT,
  status TEXT,
  collection_id TEXT,
  product_desc TEXT
);
"""

DDL_REPLY_LOG = """
CREATE TABLE IF NOT EXISTS reply_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  token TEXT,
  company TEXT,
  from_email TEXT,
  received_on TEXT,
  has_attachments INTEGER,
  parse_ok INTEGER,
  parse_json TEXT,
  collection_id TEXT,
  product_desc TEXT,
  UNIQUE(token, from_email, received_on)
);
"""

DDL_ATTACHMENT_LOG = """
CREATE TABLE IF NOT EXISTS attachment_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  token TEXT,
  msg_entryid TEXT,
  received_on TEXT,
  file_name TEXT,
  file_ext TEXT,
  file_size_bytes INTEGER,
  saved_path TEXT,
  sha256 TEXT,
  created_at TEXT
);
"""

@contextmanager
def get_conn():
    Path(DB_PATH).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    try:
        yield conn
    finally:
        conn.commit()
        conn.close()

def _ensure_column(conn: sqlite3.Connection, table: str, col: str, coltype: str):
    """Add a column if missing (SQLite ALTER TABLE ADD COLUMN is idempotent enough for our case)."""
    cur = conn.execute(f"PRAGMA table_info({table})")
    cols = [r[1] for r in cur.fetchall()]
    if col not in cols:
        try:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {coltype}")
            conn.commit()
        except Exception:
            pass

def init_db():
    with get_conn() as conn:
        c = conn.cursor()
        c.execute(DDL_MAIL_LOG)
        c.execute(DDL_REPLY_LOG)
        c.execute(DDL_ATTACHMENT_LOG)

        # Backfill columns for existing databases (compat)
        _ensure_column(conn, "reply_log", "company", "TEXT")
        _ensure_column(conn, "mail_log", "collection_id", "TEXT")
        _ensure_column(conn, "mail_log", "product_desc", "TEXT")
        _ensure_column(conn, "reply_log", "collection_id", "TEXT")
        _ensure_column(conn, "reply_log", "product_desc", "TEXT")

        conn.commit()

def insert_mail_log(row: Dict[str, Any]):
    # Provide defaults for new columns to keep callers simple
    row = {
        "collection_id": "",
        "product_desc": "",
        **row
    }
    with get_conn() as conn:
        conn.execute("""
            INSERT OR REPLACE INTO mail_log
            (email, company, token, subject, entryid, conversation_id, sent_on, status, collection_id, product_desc)
            VALUES(:email, :company, :token, :subject, :entryid, :conversation_id, :sent_on, :status, :collection_id, :product_desc)
        """, row)

def insert_reply_log(row: Dict[str, Any]):
    row = {
        "company": "",
        "collection_id": "",
        "product_desc": "",
        **row
    }
    with get_conn() as conn:
        conn.execute("""
            INSERT OR IGNORE INTO reply_log
            (token, company, from_email, received_on, has_attachments, parse_ok, parse_json, collection_id, product_desc)
            VALUES(:token, :company, :from_email, :received_on, :has_attachments, :parse_ok, :parse_json, :collection_id, :product_desc)
        """, row)

def fetch_mail_logs(limit=500):
    with get_conn() as conn:
        return conn.execute("""
            SELECT email, company, token, subject, entryid, conversation_id, sent_on, status, collection_id, product_desc
            FROM mail_log ORDER BY id DESC LIMIT ?
        """, (limit,)).fetchall()

def fetch_reply_logs(limit=500):
    with get_conn() as conn:
        return conn.execute("""
            SELECT token, company, from_email, received_on, has_attachments, parse_ok, parse_json, collection_id, product_desc
            FROM reply_log ORDER BY id DESC LIMIT ?
        """, (limit,)).fetchall()

def insert_attachment_log(row: Dict[str, Any]):
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO attachment_log
            (token, msg_entryid, received_on, file_name, file_ext, file_size_bytes, saved_path, sha256, created_at)
            VALUES(:token, :msg_entryid, :received_on, :file_name, :file_ext, :file_size_bytes, :saved_path, :sha256, :created_at)
        """, row)

def get_company_by_token(token: str) -> str:
    """Return company for a given token from mail_log (latest match), or ''."""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT company FROM mail_log WHERE token = ? ORDER BY id DESC LIMIT 1",
            (token,)
        ).fetchone()
        return (row[0] or "") if row else ""

def get_mail_meta_by_token(token: str) -> Dict[str, str]:
    """Return {company, collection_id, product_desc} for a given token."""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT company, collection_id, product_desc FROM mail_log WHERE token = ? ORDER BY id DESC LIMIT 1",
            (token,)
        ).fetchone()
        if not row:
            return {"company": "", "collection_id": "", "product_desc": ""}
        company, collection_id, product_desc = row
        return {
            "company": company or "",
            "collection_id": collection_id or "",
            "product_desc": product_desc or ""
        }
