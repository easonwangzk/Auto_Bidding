# -*- coding: utf-8 -*-
import re
import json
import uuid
from pathlib import Path
from typing import List, Tuple

import pandas as pd
import boto3
import logging
from botocore.exceptions import ClientError

from config import (
    ATTACH_BASE_DIR,
    BEDROCK_REGION, BEDROCK_AGENT_ID, BEDROCK_ALIAS_ID,
    BIDDING_MAX_PROMPT_CHARS,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

EXCEL_EXTS = {".xlsx", ".xls"}

def _sanitize_filename(name: str) -> str:
    """Sanitize a folder/file name by removing illegal characters."""
    return re.sub(r'[\\/:*?"<>|]+', "_", str(name)).strip() or "untitled"

def list_collections() -> List[str]:
    """List collection_id folders detected under ATTACH_BASE_DIR that contain Excel files."""
    out = []
    if not ATTACH_BASE_DIR.exists():
        return out
    for p in ATTACH_BASE_DIR.iterdir():
        if p.is_dir():
            has_excel = any(child.suffix.lower() in EXCEL_EXTS for child in p.iterdir() if child.is_file())
            if has_excel:
                out.append(p.name)  # folder name == safe_collection_id
    out.sort()
    return out

def list_excels_for_collection(collection_id: str) -> List[Path]:
    """Return Excel file paths under ATTACH_BASE_DIR/<safe_collection_id>/."""
    safe_collection = _sanitize_filename(collection_id)
    root = ATTACH_BASE_DIR / safe_collection
    if not root.exists():
        return []
    files = [p for p in sorted(root.iterdir()) if p.is_file() and p.suffix.lower() in EXCEL_EXTS]
    return files

def _df_to_compact_text(df: pd.DataFrame, max_rows: int = 200) -> str:
    """Turn a DataFrame into compact, readable text for LLM prompts."""
    if df is None or df.empty:
        return "(empty sheet)"
    df = df.copy()
    # Drop fully-empty columns
    df = df.dropna(axis=1, how="all")
    # Fill NaNs with blanks
    df = df.fillna("")
    # Limit rows to keep prompt smaller
    if len(df) > max_rows:
        df = df.head(max_rows)
    # Render as TSV (smaller than markdown)
    tsv = df.to_csv(index=False, sep="\t")
    return tsv

def _excel_file_to_text(path: Path) -> str:
    """Read an Excel file; convert each sheet to compact TSV blocks."""
    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        return f"# {path.name}\n(Unable to open: {e})\n"
    parts = [f"# {path.name}"]
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet, dtype=str)
            parts.append(f"## Sheet: {sheet}\n{_df_to_compact_text(df)}")
        except Exception as e:
            parts.append(f"## Sheet: {sheet}\n(Unable to read: {e})")
    return "\n".join(parts) + "\n"

def _build_prompt(collection_id: str, excel_texts: List[str], extra_instructions: str = "") -> str:
    """
    Build a concise, four-aspect prompt aligned with:
    - Product_Technical_Requirements
    - Testing_Requirements
    - Quotation
    - Additional_Claims_Features
    """
    header = (
        "You are a sourcing analyst.\n"
        "Compare supplier bids using ONLY these four aspects:\n"
        "1) Product_Technical_Requirements  — table columns: Attribute | Amazon_Requirements | Vendor_Entry.\n"
        "   - Mark Compliance per row: Yes / No / Partial, and add a short Gap_Note.\n"
        "2) Testing_Requirements            — table columns: Test_Item | Amazon_Requirements | Vendor_Entry.\n"
        "   - Mark Compliance per row: Yes / No / Partial, and add a short Gap_Note.\n"
        "3) Quotation                       — extract: Quote_Raw, Currency, Incoterm, MOQ, Lead_Time, Valid_Until.\n"
        "   - Normalize price to USD if possible; state any assumptions used.\n"
        "4) Additional_Claims_Features      — list extra certifications/materials/features (e.g., FSC wood).\n\n"
        "Steps:\n"
        "- Extract and fill the four aspects for each supplier from the provided tables.\n"
        "- Highlight gaps/differences that may affect cost, risk, or compliance.\n"
        "- Normalize prices for like-for-like comparison.\n"
        "- Rank suppliers by overall value for money (normalized price ± compliance gaps ± value-adding claims).\n\n"
        "Output:\n"
        "A) Ranked_Summary: Supplier | one-line reason.\n"
        "B) Product_Technical_Requirements_Table: Supplier | Attribute | Amazon_Requirements | Vendor_Entry | Compliance | Gap_Note.\n"
        "C) Testing_Requirements_Table:          Supplier | Test_Item | Amazon_Requirements | Vendor_Entry | Compliance | Gap_Note.\n"
        "D) Quotation_Table:                     Supplier | Quote_Raw | Currency | Incoterm | MOQ | Lead_Time | Valid_Until | Normalized_USD | Notes.\n"
        "E) Additional_Claims_Features_List:     bullets per supplier.\n"
        "F) Assumptions & Missing_Info.\n\n"
        "Rules:\n"
        "- Ignore any data not part of the four aspects above.\n"
        "- Keep concise and business-ready; use ranges when estimating; do not fabricate exact market prices.\n"
    )
    if extra_instructions:
        header += f"\nAdditional instructions:\n{extra_instructions}\n"

    body = f"\n=== COLLECTION_ID: {collection_id} ===\n\n" + "\n".join(excel_texts)
    prompt = header + "\n" + body

    if BIDDING_MAX_PROMPT_CHARS and len(prompt) > BIDDING_MAX_PROMPT_CHARS:
        prompt = prompt[:BIDDING_MAX_PROMPT_CHARS] + "\n\n[TRUNCATED]"
    return prompt

def _invoke_bedrock_agent(prompt: str, session_id: str = "") -> str:
    """Call Bedrock Agent Runtime using your streaming pattern; return completion text."""
    if not session_id:
        session_id = uuid.uuid4().hex

    client = boto3.client("bedrock-agent-runtime", region_name=BEDROCK_REGION)

    try:
        response = client.invoke_agent(
            agentId=BEDROCK_AGENT_ID,
            agentAliasId=BEDROCK_ALIAS_ID,
            enableTrace=True,
            sessionId=session_id,
            inputText=prompt,
            streamingConfigurations={
                "applyGuardrailInterval": 20,
                "streamFinalResponse": False
            }
        )
    except ClientError as e:
        logger.error("Bedrock client error: %s", e)
        raise

    completion = ""
    for event in response.get("completion", []):
        if "chunk" in event:
            chunk = event["chunk"]
            completion += chunk["bytes"].decode(errors="ignore")
        if "trace" in event:
            trace_event = event.get("trace", {})
            trace = trace_event.get("trace", {})
            for k, v in trace.items():
                logger.info("%s: %s", k, v)
    return completion.strip()

def compare_bids(collection_id: str, extra_instructions: str = "") -> Tuple[str, List[Path]]:
    """
    Load Excel files for the given collection, build a prompt and get Bedrock output.
    Returns (model_output, file_list).
    """
    files = list_excels_for_collection(collection_id)
    if not files:
        return f"No Excel files found for collection: {collection_id}", []

    excel_texts = [_excel_file_to_text(p) for p in files]
    prompt = _build_prompt(collection_id, excel_texts, extra_instructions=extra_instructions)
    output = _invoke_bedrock_agent(prompt)
    return output, files

