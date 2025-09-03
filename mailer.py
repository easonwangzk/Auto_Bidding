# -*- coding: utf-8 -*-
import time
import uuid
import threading
from datetime import datetime, timezone
import pandas as pd
from jinja2 import Template
import win32com.client as win32
import pythoncom
from pathlib import Path
from typing import List, Optional

from config import (
    DEFAULT_CC, SUBJECT_BASE, SEND_THROTTLE,
    OL_MAILITEM, OL_FOLDER_INBOX, OL_FOLDER_SENTMAIL,
    USER_PROP_TOKEN, TOKEN_PREFIX
)
from db import insert_mail_log

_thread_state = threading.local()

def _ensure_com_initialized():
    """Initialize COM for the current thread; returns a cleanup function if needed."""
    if getattr(_thread_state, "com_inited", False):
        return None
    pythoncom.CoInitialize()
    _thread_state.com_inited = True
    def _cleanup():
        try:
            pythoncom.CoUninitialize()
        finally:
            _thread_state.com_inited = False
    return _cleanup

def _ensure_session(app):
    """Ensure the MAPI session is available."""
    ns = app.GetNamespace("MAPI")
    try:
        _ = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    except Exception:
        ns.Logon("", "", False, False)
        _ = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    return ns

def _load_template(path: str) -> Template:
    """Load an HTML email template from a file."""
    with open(path, "r", encoding="utf-8") as f:
        return Template(f.read())

def _safe_str(v):
    """Return trimmed string or empty if NaN/None."""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip() if v is not None else ""

def _add_attachments(mail, attach_paths: Optional[List[Path]]):
    """Attach files to the Outlook mail item if paths are provided."""
    if not attach_paths:
        return
    for p in attach_paths:
        try:
            p = Path(p)
            if p.exists() and p.is_file():
                # Keep original filename; do not rename
                mail.Attachments.Add(str(p))
        except Exception:
            # Best-effort: skip problematic files without failing the whole send
            pass

def send_one(app, ns,
             to_email: str,
             company: str,
             collection_id: str,
             product_desc: str,
             html_template: Template,
             attach_paths: Optional[List[Path]] = None):
    """
    Send a single email and log it into mail_log (ensures COM initialization for the current thread).
    Optionally attach files listed in attach_paths.
    """
    cleanup = _ensure_com_initialized()
    try:
        token = f"{TOKEN_PREFIX}{uuid.uuid4().hex[:8].upper()}"
        subject = f"[{token}] {SUBJECT_BASE}"

        # Render template; extra variables are safe even if template does not use them.
        html_body = html_template.render(
            supplier_name=company or "Valued Supplier",
            token=token,
            collection_id=collection_id,
            product_desc=product_desc
        )

        mail = app.CreateItem(OL_MAILITEM)
        mail.To = to_email
        if DEFAULT_CC:
            mail.CC = DEFAULT_CC
        mail.Subject = subject
        mail.HTMLBody = html_body

        # Custom user properties (1 = olText). Names avoid special chars.
        try:
            up = mail.UserProperties
            up.Add(USER_PROP_TOKEN, 1, True).Value = token
            up.Add("CollectionID", 1, True).Value = collection_id
            up.Add("ProductDescription", 1, True).Value = product_desc
        except Exception:
            # User properties are optional; ignore failures.
            pass

        # Add fixed attachments BEFORE sending
        _add_attachments(mail, attach_paths)

        mail.Save()
        mail.Send()

        # Retrieve identifiers from Sent Items
        sent_on = datetime.now(timezone.utc).isoformat()
        entryid, conversation_id = None, None
        try:
            sent_folder = ns.GetDefaultFolder(OL_FOLDER_SENTMAIL)
            items = sent_folder.Items
            items.Sort("[SentOn]", True)
            for i in range(1, min(50, items.Count) + 1):
                it = items.Item(i)
                if token in (getattr(it, "Subject", "") or ""):
                    entryid = getattr(it, "EntryID", None)
                    conversation_id = getattr(it, "ConversationID", None)
                    sent_on = str(getattr(it, "SentOn", sent_on))
                    break
        except Exception:
            pass

        # Write to mail_log including collection_id / product_desc
        insert_mail_log({
            "email": to_email,
            "company": company or "",
            "token": token,
            "subject": subject,
            "entryid": entryid or "",
            "conversation_id": conversation_id or "",
            "sent_on": sent_on,
            "status": "SENT",
            "collection_id": collection_id or "",
            "product_desc": product_desc or ""
        })

        return token, subject
    finally:
        if callable(cleanup):
            cleanup()

def bulk_send(contacts_df: pd.DataFrame, template_path: str, attach_paths: Optional[List[Path]] = None):
    """
    Expected columns in contacts_df:
      - Email (required)
      - Company Name (optional)
      - Collection ID (optional)
      - Product description (optional)
    Optionally pass a list of attachment paths to include in every email.
    """
    cleanup = _ensure_com_initialized()
    try:
        app = win32.Dispatch("Outlook.Application")
        ns = _ensure_session(app)
        tpl = _load_template(template_path)

        results = []
        for _, row in contacts_df.iterrows():
            email = _safe_str(row.get("Email"))
            if not email:
                continue

            company = _safe_str(row.get("Company Name")) or "Valued Supplier"
            collection_id = _safe_str(row.get("Collection ID"))
            product_desc = _safe_str(row.get("Product description"))

            token, subject = send_one(
                app, ns, email, company, collection_id, product_desc, tpl,
                attach_paths=attach_paths
            )
            results.append({
                "email": email,
                "company": company,
                "collection_id": collection_id,
                "product_desc": product_desc,
                "token": token,
                "subject": subject
            })
            # Throttle to avoid security prompts/rate issues
            time.sleep(SEND_THROTTLE)

        return results
    finally:
        if callable(cleanup):
            cleanup()
