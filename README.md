Auto Bidding Platform

A Windows-based automation platform designed to support the auto bidding workflow for sourcing teams.
It integrates Excel-based supplier/product data comparison, automated email outreach with tracking tokens, reply polling, and attachment logging — all via a simple web UI.
Built on Microsoft Outlook COM API and Streamlit, it enables batch, trackable supplier communication, and centralizes bid data for evaluation.

====================================
Features
====================================

1. Supplier & Product Data Comparison
- Reads Excel submissions from multiple suppliers.
- Compares only the specified attributes relevant to the product category (e.g., material, dimensions, load capacity, certifications).
- Identifies attribute differences that may impact cost or risk.
- Normalizes prices to like-for-like baselines and ranks bids by overall value for money.

2. Batch Email Sending (Outlook COM Automation)
- Reads contact list from contacts.xlsx (required) and optional product list from products.xlsx.
- Automatically appends tracking token (e.g., [ABA#XXXXXXXX]) to subject lines.
- Stores the token in a custom Outlook user property (ABATrackingToken) for reply matching.

3. Inbox Polling & Reply Capture
- Scans configured Outlook folders (Inbox + subfolders).
- Matches incoming replies by tracking token in subject or body.
- Logs sender, received time, and whether attachments are present.

4. Attachment Management
- Saves and logs attachments from matched replies.
- Enforces file size limits and avoids filename conflicts automatically.

5. Web Interface (Streamlit)
- Tab 1: Send Emails — upload contacts/products and send batch messages.
- Tab 2: Poll Inbox — trigger inbox scans for replies.
- Tab 3: View Logs — browse sent mail log and reply log.

====================================
Prerequisites
====================================
- Windows with Microsoft Outlook installed (same bitness as Python, usually 64-bit).
- Outlook account configured and logged in.
- Python 3.9+ (64-bit recommended).

Install dependencies:
pip install -r requirements.txt

Key dependencies:
- pywin32 — COM automation for Outlook
- streamlit — Web UI
- pandas — Excel processing
- pythoncom — COM thread management

====================================
Quick Start
====================================
1. Verify Outlook COM access:
   python -c "import win32com.client; print('pywin32 ok')"

2. Launch Web Interface
   $env:AWS_ACCESS_KEY_ID="AKIARM3JAK3JIVBNBVGL"
   $env:AWS_SECRET_ACCESS_KEY="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"
   $env:AWS_DEFAULT_REGION="us-east-2"
   streamlit run app.py

====================================
Usage Guide
====================================

Tab 1: Send Emails
- Upload contacts.xlsx (contacts sheet):
  - Email — recipient address (required)
  - Company Name — optional
  - Other fields used in email template rendering
- Optionally upload products.xlsx (products sheet):
  - ASIN — matched to recipients for personalized info
- Click Send to:
  - Insert tracking token in subject
  - Save token to ABATrackingToken property
  - Send all emails (with throttling)

Tab 2: Poll Inbox
- Click Poll Inbox Now to:
  - Scan folders in config.SCAN_FOLDERS (or Inbox by default)
  - Detect replies containing tracking tokens in subject/body
  - Log reply details to reply_log (attachments logged separately)

Tab 3: View Logs
- mail_log — sent email records (tokens, recipients, status)
- reply_log — reply metadata (sender, received time, attachments)

====================================
Configuration (config.py)
====================================
- TOKEN_PREFIX — tracking token prefix, e.g., ABA#
- POLL_LOOKBACK_MINUTES — how far back to scan for replies
- POLL_MAX_SCAN — max items per folder scan
- SCAN_FOLDERS — folders to scan (e.g., ["Inbox", "Inbox/Attachments"])
- MAX_ATTACH_SIZE_MB — skip saving attachments over this size
- ATTACH_BASE_DIR — base directory for saving attachments

====================================
Notes & Tips
====================================
- First run may trigger Outlook programmatic access prompts:
  1. File → Options → Trust Center → Trust Center Settings → Programmatic Access
  2. Allow programmatic access if policy permits
- Token format: [ABA#XXXXXXXX] (8-char uppercase alphanumeric)
- Reply matching:
  1. Subject token match
  2. Body token match (fallback)
- Attachment saving skips inline images and corrupted files, logs all attempts
- Throttling, polling, and folder scan list adjustable in config.py

====================================
Developer Notes
====================================
- COM calls wrapped with com_apartment() for thread-safe Outlook automation
- Folder scanning supports nested paths like "Inbox/Subfolder1"
- Logging functions:
  - insert_reply_log() — reply metadata
  - insert_attachment_log() — attachment metadata
- Supports manual polling and scheduled polling (Windows Task Scheduler)

====================================
License
====================================
MIT License — Use at your own risk, especially with programmatic email sending in Outlook.