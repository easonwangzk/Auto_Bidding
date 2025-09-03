# -*- coding: utf-8 -*-
import re
import hashlib
import os
import json
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from typing import Iterable, List

import pythoncom
import win32com.client as win32

from config import (
    OL_FOLDER_INBOX,
    USER_PROP_TOKEN,
    TOKEN_PREFIX,
    POLL_LOOKBACK_MINUTES,
    POLL_MAX_SCAN,
    ATTACH_BASE_DIR,
    MAX_ATTACH_SIZE_MB,
)

try:
    from config import SCAN_FOLDERS  # type: ignore
except Exception:
    SCAN_FOLDERS = []

# NOTE: now pulling full meta by token (company, collection_id, product_desc)
from db import insert_reply_log, insert_attachment_log, get_mail_meta_by_token

TOKEN_RE = re.compile(r"\[{}[A-Z0-9]{{8}}\]".format(re.escape(TOKEN_PREFIX)), re.I)
_RPC_E_CHANGED_MODE = -2147417850  # 0x80010106

# Limit stored body size to avoid huge blobs (set 0/None to disable)
MAX_BODY_TEXT_CHARS = 100000


@contextmanager
def com_apartment():
    """Ensure COM apartment-threaded initialization for the current context."""
    inited_here = False
    try:
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            inited_here = True
        except pythoncom.com_error as e:
            if getattr(e, "hresult", None) == _RPC_E_CHANGED_MODE:
                pass
            else:
                raise
        yield
    finally:
        if inited_here:
            try:
                pythoncom.CoUninitialize()
            except pythoncom.com_error:
                pass


def _ensure_session(app):
    """Ensure an active Outlook MAPI session."""
    ns = app.GetNamespace("MAPI")
    try:
        _ = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    except Exception:
        ns.Logon("", "", False, False)
        _ = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    return ns


def _cutoff_dt():
    """Return the datetime cutoff for scanning messages."""
    return datetime.now() - timedelta(minutes=POLL_LOOKBACK_MINUTES)


def _sender_email(item) -> str:
    """Extract sender's email address from a mail item."""
    try:
        sender = getattr(item, "Sender", None)
        if sender is not None:
            try:
                exu = sender.GetExchangeUser()
                if exu and exu.PrimarySmtpAddress:
                    return exu.PrimarySmtpAddress
            except Exception:
                pass
            try:
                addr = sender.Address
                if addr:
                    return addr
            except Exception:
                pass
    except Exception:
        pass
    try:
        return item.SenderEmailAddress or ""
    except Exception:
        return ""


def _resolve_folder_by_path(root_folder, path_parts: List[str]):
    """Traverse folder hierarchy by path parts to resolve the folder object."""
    folder = root_folder
    for part in path_parts[1:]:
        found = None
        try:
            for f in folder.Folders:
                if str(f.Name).strip().lower() == str(part).strip().lower():
                    found = f
                    break
        except Exception:
            pass
        if not found:
            return None
        folder = found
    return folder


def _iter_configured_folders(ns) -> Iterable:
    """Return the list of folders specified in SCAN_FOLDERS."""
    if not SCAN_FOLDERS:
        return []
    resolved = []
    for store in ns.Stores:
        try:
            inbox = store.GetDefaultFolder(OL_FOLDER_INBOX)
        except Exception:
            continue
        for raw in SCAN_FOLDERS:
            path = str(raw).strip()
            if not path:
                continue
            parts = [p for p in path.split("/") if p]
            if not parts or str(parts[0]).strip().lower() != "inbox":
                continue
            if len(parts) == 1:
                resolved.append(inbox)
            else:
                f = _resolve_folder_by_path(inbox, parts)
                if f:
                    resolved.append(f)
    uniq = {}
    for f in resolved:
        try:
            k = getattr(f, "EntryID", None) or id(f)
            uniq[k] = f
        except Exception:
            pass
    return list(uniq.values())


def _iter_default_folders(ns) -> Iterable:
    """Return Inbox and its first-level subfolders for all accounts."""
    for store in ns.Stores:
        try:
            inbox = store.GetDefaultFolder(OL_FOLDER_INBOX)
        except Exception:
            continue
        yield inbox
        try:
            for f in inbox.Folders:
                yield f
        except Exception:
            pass


def _sanitize_filename(name: str) -> str:
    """Remove invalid characters from a file name."""
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "attachment"


def _unique_path(dest_dir, base_name):
    """If a file with the same name exists, append (1), (2), etc."""
    name, ext = os.path.splitext(base_name)
    p = os.path.join(dest_dir, base_name)
    idx = 1
    while os.path.exists(p):
        p = os.path.join(dest_dir, f"{name}({idx}){ext}")
        idx += 1
    return p


def _sha256_of_file(path: str) -> str:
    """Compute the SHA256 hash of a file."""
    try:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return ""


def _save_attachments(company: str, token: str, collection_id: str, msg, msg_entryid: str, received_on: str):
    """
    Save all attachments and log them.
    Folder: ATTACH_BASE_DIR / <safe_collection_id>
    Filename: {safe_company}_{safe_token}{ext}
    """
    try:
        atts = getattr(msg, "Attachments", None)
        if atts is None or atts.Count == 0:
            return 0
    except Exception:
        return 0

    # sanitize parts
    safe_company = _sanitize_filename(company or "")
    safe_token = _sanitize_filename(token or "")
    safe_collection = _sanitize_filename(collection_id or "uncategorized")

    # target folder grouped by collection_id
    dest_dir = ATTACH_BASE_DIR / safe_collection
    dest_dir.mkdir(parents=True, exist_ok=True)

    saved = 0
    for i in range(1, atts.Count + 1):
        try:
            att = atts.Item(i)
        except Exception:
            continue

        # original filename & size
        try:
            orig_fname = _sanitize_filename(att.FileName or f"attachment_{i}")
        except Exception:
            orig_fname = f"attachment_{i}"

        fsize = 0
        try:
            fsize = int(getattr(att, "Size", 0))
        except Exception:
            pass

        # build target filename: {safe_company}_{safe_token}{ext}
        ext = os.path.splitext(orig_fname)[1].lower()
        name_core = f"{safe_company}_{safe_token}" if safe_company else safe_token
        target_fname = f"{name_core}{ext}"

        # Optional size limit — log metadata only
        if MAX_ATTACH_SIZE_MB and fsize > MAX_ATTACH_SIZE_MB * 1024 * 1024:
            insert_attachment_log({
                "token": token,
                "msg_entryid": msg_entryid,
                "received_on": received_on,
                "file_name": target_fname,
                "file_ext": ext,
                "file_size_bytes": fsize,
                "saved_path": "",
                "sha256": "",
                "created_at": datetime.now(timezone.utc).isoformat(),
            })
            continue

        # ensure unique path within collection folder
        save_path = _unique_path(str(dest_dir), target_fname)
        try:
            att.SaveAsFile(save_path)
        except Exception:
            # saving failed — log metadata only
            insert_attachment_log({
                "token": token,
                "msg_entryid": msg_entryid,
                "received_on": received_on,
                "file_name": target_fname,
                "file_ext": ext,
                "file_size_bytes": fsize,
                "saved_path": "",
                "sha256": "",
                "created_at": datetime.now(timezone.utc).isoformat(),
            })
            continue

        sha = _sha256_of_file(save_path)
        insert_attachment_log({
            "token": token,
            "msg_entryid": msg_entryid,
            "received_on": received_on,
            "file_name": target_fname,
            "file_ext": ext,
            "file_size_bytes": fsize,
            "saved_path": save_path,
            "sha256": sha,
            "created_at": datetime.now(timezone.utc).isoformat(),
        })
        saved += 1

    return saved


def poll_inbox():
    """Scan configured Outlook folders for matching tokens and save attachments."""
    hits = 0
    scanned = 0
    cutoff = _cutoff_dt()

    with com_apartment():
        app = win32.Dispatch("Outlook.Application")
        ns = _ensure_session(app)

        folders = _iter_configured_folders(ns)
        if not folders:
            folders = list(_iter_default_folders(ns))

        for folder in folders:
            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)  # newest first
            except Exception:
                continue

            count = min(POLL_MAX_SCAN, getattr(items, "Count", 0))
            for i in range(1, count + 1):
                try:
                    it = items.Item(i)
                except Exception:
                    continue

                # Optional cutoff (noop here; keep original behavior)
                try:
                    rc = getattr(it, "ReceivedTime", None)
                    if rc and hasattr(rc, "year"):
                        pass
                except Exception:
                    pass

                subj = (getattr(it, "Subject", "") or "")
                body_plain = (getattr(it, "Body", "") or "")  # Outlook's plain text body

                token = None
                m = TOKEN_RE.search(subj) or TOKEN_RE.search(body_plain)
                if m:
                    token = m.group(0)[1:-1]
                if not token:
                    # Fallback: try HTMLBody
                    try:
                        body_html = getattr(it, "HTMLBody", "") or ""
                    except Exception:
                        body_html = ""
                    m2 = TOKEN_RE.search(body_html)
                    if m2:
                        token = m2.group(0)[1:-1]

                if not token:
                    continue

                # Lookup metadata from mail_log by token
                try:
                    meta = get_mail_meta_by_token(token) or {}
                except Exception:
                    meta = {}
                company = meta.get("company", "") or ""
                collection_id = meta.get("collection_id", "") or ""
                product_desc = meta.get("product_desc", "") or ""

                from_email = _sender_email(it)
                try:
                    has_attachments = int(getattr(it, "Attachments", None) is not None and it.Attachments.Count > 0)
                except Exception:
                    has_attachments = 0
                try:
                    received_on = str(getattr(it, "ReceivedTime", ""))
                except Exception:
                    received_on = datetime.now(timezone.utc).isoformat()

                # Prepare parse_json: store subject + body_text
                body_text = body_plain
                if not body_text:
                    try:
                        body_html = getattr(it, "HTMLBody", "") or ""
                    except Exception:
                        body_html = ""
                    if body_html:
                        body_text = re.sub(r"<[^>]+>", " ", body_html)
                        body_text = re.sub(r"\s+\n", "\n", body_text)
                        body_text = re.sub(r"[ \t]{2,}", " ", body_text)

                if MAX_BODY_TEXT_CHARS and isinstance(body_text, str) and len(body_text) > MAX_BODY_TEXT_CHARS:
                    body_text = body_text[:MAX_BODY_TEXT_CHARS]

                parse_json_payload = json.dumps(
                    {"subject": subj, "body_text": body_text},
                    ensure_ascii=False
                )

                # Insert reply log (with company + collection_id + product_desc)
                insert_reply_log({
                    "token": token,
                    "company": company,
                    "from_email": from_email,
                    "received_on": received_on,
                    "has_attachments": has_attachments,
                    "parse_ok": 0,
                    "parse_json": parse_json_payload,
                    "collection_id": collection_id,
                    "product_desc": product_desc
                })
                hits += 1
                scanned += 1

                # Save attachments grouped by collection_id, named {company}_{token}{ext}
                if has_attachments:
                    try:
                        msg_entryid = getattr(it, "EntryID", "") or ""
                    except Exception:
                        msg_entryid = ""
                    _save_attachments(company, token, collection_id, it, msg_entryid, received_on)

    return {"scanned": scanned, "matched": hits}
