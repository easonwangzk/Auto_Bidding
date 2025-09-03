# -*- coding: utf-8 -*-
from pathlib import Path

# Database file
DB_PATH = Path(__file__).parent / "mail.db"

# Default email settings
DEFAULT_CC = ""   # Can be an empty string
SUBJECT_BASE = "Invitation to Partner with Amazon Private Brands"

# Send throttle (seconds)
SEND_THROTTLE = 0.8

# Inbox polling parameters
POLL_LOOKBACK_MINUTES = 7 * 24 * 60   # Last 7 days
POLL_MAX_SCAN = 400

# Outlook constants
OL_MAILITEM = 0
OL_FOLDER_INBOX = 6
OL_FOLDER_SENTMAIL = 5

# Custom user property name (cannot contain [ ] _ #)
USER_PROP_TOKEN = "ABATrackingToken"

# Token prefix in subject
TOKEN_PREFIX = "ABA#" 

# List of folders to scan (empty means scan all accounts' Inbox + one-level subfolders)
SCAN_FOLDERS = [
    "Inbox",                 # Main Inbox
    "Inbox/External",        # One-level subfolder
    # "Inbox/Suppliers/Quotes" # Multi-level subfolder example
]

# Attachment base directory (subfolders per token)
ATTACH_BASE_DIR = Path(__file__).parent / "attachments"
ATTACH_BASE_DIR.mkdir(parents=True, exist_ok=True)

# Optional: Maximum size for a single attachment (MB), 0 means no limit
MAX_ATTACH_SIZE_MB = 50

# Attach one or more fixed files to every outgoing email.
ATTACH_FILES = [
    Path(__file__).parent / "data" / "supplier_form.xlsx",  
]

# --- Bedrock / Bidding_Comparison ---
BEDROCK_REGION = "us-east-2"
BEDROCK_AGENT_ID = "1NSQC29UJR"
BEDROCK_ALIAS_ID = "8UYSFGE2GE"

# Prompt size guard (characters). Prevents huge payloads.
BIDDING_MAX_PROMPT_CHARS = 120_000


