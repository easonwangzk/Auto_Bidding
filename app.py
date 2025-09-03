"""
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path

from db import init_db, fetch_mail_logs, fetch_reply_logs
from mailer import bulk_send
from poller import poll_inbox
from config import ATTACH_FILES  # fixed attachment paths (list of Path or str)

# NEW: bidding comparison helpers
from bidding_comparison import (
    list_collections,
    list_excels_for_collection, 
    compare_bids,
)

st.set_page_config(page_title="Auto Bidding Platform", layout="wide")
st.title("Auto Bidding Platform")

init_db()

with st.sidebar:
    st.markdown("### Quick Actions")
    if st.button("Poll Inbox Now"):
        result = poll_inbox()
        st.success(f"Scanned: {result['scanned']}, Matched: {result['matched']}")
    st.markdown("---")
    st.caption("Make sure Outlook is open and logged in.")

tab1, tab2, tab3, tab4 = st.tabs(["① Upload Contacts", "② Send Emails", "③ Logs", "④ Bidding Comparison"])

# --------------------
# Tab 1: Upload Contacts
# --------------------
with tab1:
    st.subheader("Upload contacts.xlsx (sheet: contacts)")
    cfile = st.file_uploader("contacts.xlsx", type=["xlsx"])

    contacts_df = pd.DataFrame()
    if cfile:
        contacts_df = pd.read_excel(cfile, sheet_name="contacts")
        contacts_df.columns = contacts_df.columns.str.strip()
        st.write("Contacts preview:", contacts_df.head(20))
        st.info("Required: Email; Optional: Company Name, Collection ID, Product description")
        # Lightweight validation: require Email column
        if "Email" not in contacts_df.columns:
            st.error("Contacts is missing the required column: Email")
    else:
        st.warning("Please upload contacts.xlsx with sheet name 'contacts'.")

    st.session_state["contacts_df"] = contacts_df

# --------------------
# Tab 2: Send Emails
# --------------------
with tab2:
    st.subheader("Send Emails")
    template_path = Path("templates/email_template.html")
    if not template_path.exists():
        st.error("Missing templates/email_template.html")
    else:
        st.code(f"Using template: {template_path}")

    # Fixed attachments (Option A) preview
    st.markdown("**Fixed attachments to include in every email (from config):**")
    attach_existing = []
    attach_missing = []
    for p in ATTACH_FILES:
        p = Path(p)
        if p.exists() and p.is_file():
            attach_existing.append(p)
        else:
            attach_missing.append(p)

    if attach_existing:
        for p in attach_existing:
            st.write(f" {p}")
    else:
        st.write("— None found —")

    if attach_missing:
        st.warning("Some configured attachment paths do not exist:")
        for p in attach_missing:
            st.write(f" {p}")

    if st.button("Send (Small Batch First)"):
        contacts_df = st.session_state.get("contacts_df", pd.DataFrame())
        if contacts_df.empty:
            st.error("Contacts is empty.")
        elif "Email" not in contacts_df.columns:
            st.error("Contacts is missing the required column: Email")
        else:
            with st.spinner("Sending..."):
                # Only pass existing files (Outlook requires real paths)
                results = bulk_send(
                    contacts_df,
                    str(template_path),
                    attach_paths=attach_existing
                )
            st.success(f"Done. Sent: {len(results)}")
            st.write(pd.DataFrame(results))

# --------------------
# Tab 3: Logs
# --------------------
with tab3:
    st.subheader("Mail Logs")
    mlogs = fetch_mail_logs(500)
    if mlogs:
        # db.py returns 10 columns including collection_id and product_desc
        mdf = pd.DataFrame(
            mlogs,
            columns=[
                "email", "company", "token", "subject", "entryid",
                "conversation_id", "sent_on", "status", "collection_id", "product_desc"
            ]
        )
        st.dataframe(mdf, use_container_width=True, height=320)
    else:
        st.info("No mail logs yet.")

    st.subheader("Reply Logs")
    rlogs = fetch_reply_logs(500)
    if rlogs:
        # db.py returns 9 columns including collection_id and product_desc
        rdf = pd.DataFrame(
            rlogs,
            columns=[
                "token", "company", "from_email", "received_on",
                "has_attachments", "parse_ok", "parse_json",
                "collection_id", "product_desc"
            ]
        )
        st.dataframe(rdf, use_container_width=True, height=320)
    else:
        st.info("No replies captured yet.")

# --------------------
# Tab 4: Bidding Comparison
# --------------------
with tab4:
    st.subheader("Bidding Comparison")

    collections = list_collections()
    if not collections:
        st.info("No collections with Excel attachments found yet. Poll inbox and ensure attachments are saved.")
    else:
        cid = st.selectbox("Select a collection_id", collections, index=0)

        files = list_excels_for_collection(cid)
        if files:
            st.write("Files to compare:")
            for p in files:
                st.write(f"• {p.name}")
        else:
            st.warning("No Excel files under this collection.")

        extra = st.text_area(
            "Optional: Additional instructions for the model",
            placeholder="Example: prioritize GRS ≥50% over Recycled 100 if price difference > 10% ..."
        )

        if st.button("Run Comparison"):
            with st.spinner("Calling Bedrock…"):
                try:
                    output, used_files = compare_bids(cid, extra_instructions=extra.strip())
                    st.success("Comparison complete.")
                    st.markdown("**Model Output:**")
                    st.text_area("Result", value=output, height=400)
                except Exception as e:
                    st.error(f"Bedrock comparison failed: {e}")

"""
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from pathlib import Path
import hashlib

from db import init_db, fetch_mail_logs, fetch_reply_logs
from mailer import bulk_send
from poller import poll_inbox
from config import ATTACH_FILES  # fixed attachment paths (list of Path or str)

# NEW: bidding comparison helpers
from bidding_comparison import (
    list_collections,
    list_excels_for_collection,
    compare_bids,
)

# =========================
# Utilities
# =========================
def make_key(prefix: str, *parts) -> str:
    """
    Create a globally-unique, Streamlit-safe key from arbitrary parts.
    We hash the concatenated parts to avoid collisions and long strings.
    """
    raw = prefix + "||" + "||".join(str(p) for p in parts)
    digest = hashlib.md5(raw.encode("utf-8")).hexdigest()[:12]
    return f"{prefix}_{digest}"

@st.cache_data(show_spinner=False)
def list_files_for_collection(cid: str) -> list[str]:
    """Cached absolute paths of Excel files for a given collection_id."""
    return [str(p) for p in list_excels_for_collection(cid)]

@st.cache_data(show_spinner=False)
def get_sheet_names_for_file(path_str: str, mtime: float) -> list[str]:
    """Return sheet names for an Excel file. mtime included to refresh cache if file changes."""
    xls = pd.ExcelFile(path_str)
    return list(xls.sheet_names)

@st.cache_data(show_spinner=False)
def load_excel_sheet(path_str: str, sheet_name: str, mtime: float) -> pd.DataFrame:
    """Load a specific sheet as strings. mtime included to refresh cache if file changes."""
    return pd.read_excel(path_str, sheet_name=sheet_name, dtype=str)

# =========================
# App Bootstrap
# =========================
st.set_page_config(page_title="Auto Bidding Platform", layout="wide")
st.title("Auto Bidding Platform")
init_db()

with st.sidebar:
    st.markdown("### Quick Actions")
    if st.button("Poll Inbox Now", key=make_key("btn_poll_inbox")):
        result = poll_inbox()
        st.success(f"Scanned: {result['scanned']}, Matched: {result['matched']}")
    st.markdown("---")
    st.caption("Make sure Outlook is open and logged in.")

tab1, tab2, tab3, tab4 = st.tabs(["① Upload Contacts", "② Send Emails", "③ Logs", "④ Bidding Comparison"])

# =========================
# Tab 1: Upload Contacts
# =========================
with tab1:
    st.subheader("Upload contacts.xlsx (sheet: contacts)")
    cfile = st.file_uploader("contacts.xlsx", type=["xlsx"], key=make_key("upl_contacts"))

    contacts_df = pd.DataFrame()
    if cfile:
        contacts_df = pd.read_excel(cfile, sheet_name="contacts")
        contacts_df.columns = contacts_df.columns.str.strip()
        st.write("Contacts preview:", contacts_df.head(20))
        st.info("Required: Email; Optional: Company Name, Collection ID, Product description")
        if "Email" not in contacts_df.columns:
            st.error("Contacts is missing the required column: Email")
    else:
        st.warning("Please upload contacts.xlsx with sheet name 'contacts'.")

    st.session_state["contacts_df"] = contacts_df

# =========================
# Tab 2: Send Emails
# =========================
with tab2:
    st.subheader("Send Emails")
    template_path = Path("templates/email_template.html")
    if not template_path.exists():
        st.error("Missing templates/email_template.html")
    else:
        st.code(f"Using template: {template_path}")

    # Fixed attachments preview
    st.markdown("**Fixed attachments to include in every email (from config):**")
    attach_existing, attach_missing = [], []
    for p in ATTACH_FILES:
        p = Path(p)
        (attach_existing if (p.exists() and p.is_file()) else attach_missing).append(p)

    if attach_existing:
        for p in attach_existing:
            st.write(f"{p}")
    else:
        st.write("— None found —")

    if attach_missing:
        st.warning("Some configured attachment paths do not exist:")
        for p in attach_missing:
            st.write(f"{p}")

    if st.button("Send (Small Batch First)", key=make_key("btn_send_small")):
        contacts_df = st.session_state.get("contacts_df", pd.DataFrame())
        if contacts_df.empty:
            st.error("Contacts is empty.")
        elif "Email" not in contacts_df.columns:
            st.error("Contacts is missing the required column: Email")
        else:
            with st.spinner("Sending..."):
                results = bulk_send(
                    contacts_df,
                    str(template_path),
                    attach_paths=attach_existing
                )
            st.success(f"Done. Sent: {len(results)}")
            st.write(pd.DataFrame(results))

# =========================
# Tab 3: Logs
# =========================
with tab3:
    st.subheader("Mail Logs")
    mlogs = fetch_mail_logs(500)
    if mlogs:
        mdf = pd.DataFrame(
            mlogs,
            columns=[
                "email", "company", "token", "subject", "entryid",
                "conversation_id", "sent_on", "status", "collection_id", "product_desc"
            ]
        )
        st.dataframe(mdf, use_container_width=True, height=320)
    else:
        st.info("No mail logs yet.")

    st.subheader("Reply Logs")
    rlogs = fetch_reply_logs(500)
    if rlogs:
        rdf = pd.DataFrame(
            rlogs,
            columns=[
                "token", "company", "from_email", "received_on",
                "has_attachments", "parse_ok", "parse_json",
                "collection_id", "product_desc"
            ]
        )
        st.dataframe(rdf, use_container_width=True, height=320)
    else:
        st.info("No replies captured yet.")

# =========================
# Tab 4: Bidding Comparison
# =========================
with tab4:
    st.subheader("Bidding Comparison")

    collections_all = list_collections()
    if not collections_all:
        st.info("No collections with Excel attachments found yet. Poll inbox and ensure attachments are saved.")
        st.stop()

    # Single-collection selector for LLM & per-file panels
    cid = st.selectbox(
        "Select a collection_id (for LLM & per-file panels)",
        collections_all,
        index=0,
        key=make_key("sel_collection_single")
    )

    files = list_excels_for_collection(cid)
    if files:
        st.write("Files detected under this collection:")
        for p in files:
            st.write(f"• {p.name}")
    else:
        st.warning("No Excel files under this collection.")

    # Three sub-tabs
    sub_tab_compare, sub_tab_view, sub_tab_browser = st.tabs(
        ["Compare (LLM)", "Explore (Per-file Panels)", "Explore (Collections Browser)"]
    )

    # ---- Sub-tab 1: Compare (LLM) ----
    with sub_tab_compare:
        extra = st.text_area(
            "Optional: Additional instructions for the model",
            placeholder="Example: prioritize GRS ≥50% over Recycled 100 if price difference > 10% ...",
            key=make_key("txt_extra_llm", cid)
        )

        if st.button("Run Comparison", key=make_key("btn_run_llm", cid)):
            with st.spinner("Calling Bedrock…"):
                try:
                    output, used_files = compare_bids(cid, extra_instructions=extra.strip())
                    st.success("Comparison complete.")
                    st.markdown("**Model Output:**")
                    st.text_area("Result", value=output, height=400, key=make_key("txt_llm_result", cid))
                except Exception as e:
                    st.error(f"Bedrock comparison failed: {e}")

    # ---- Sub-tab 2: Explore (Per-file Panels) for selected collection ----
    with sub_tab_view:
        st.caption("Pick multiple files from the selected collection; each file renders in its own panel (no merge).")

        if not files:
            st.stop()

        file_names = [p.name for p in files]
        default_pick = file_names[:min(3, len(file_names))]
        selected_names = st.multiselect(
            "Files to display",
            options=file_names,
            default=default_pick,
            key=make_key("ms_files_single_collection", cid)
        )
        if not selected_names:
            st.info("Select at least one file to display.")
            st.stop()

        name_to_path = {p.name: str(p) for p in files}

        # Render a panel for each selected file
        for idx, nm in enumerate(selected_names):
            pstr = name_to_path.get(nm)
            if not pstr or not Path(pstr).exists():
                st.warning(f"File not found: {nm}")
                continue

            mtime = Path(pstr).stat().st_mtime
            try:
                sheets = get_sheet_names_for_file(pstr, mtime)
            except Exception as e:
                st.warning(f"Unable to open {nm}: {e}")
                continue
            if not sheets:
                st.warning(f"No readable sheets in {nm}.")
                continue

            with st.expander(f"📄 {nm}", expanded=False):
                sheet_key = make_key("sel_sheet_single", cid, nm, idx)
                sheet = st.selectbox(
                    f"Sheet to view ({nm})",
                    options=sheets,
                    index=0,
                    key=sheet_key
                )

                try:
                    df = load_excel_sheet(pstr, sheet, mtime).copy()
                except Exception as e:
                    st.warning(f"Unable to read {nm}/{sheet}: {e}")
                    continue

                if df is None or df.empty:
                    st.info("This sheet has no rows.")
                    continue

                all_cols = list(df.columns)
                cols_key = make_key("ms_cols_single", cid, nm, idx)
                picked = st.multiselect(
                    f"Columns to show ({nm})",
                    options=all_cols,
                    default=all_cols,
                    key=cols_key
                )
                final_cols = picked if picked else all_cols

                cap_key = make_key("num_cap_single", cid, nm, idx)
                cap = st.number_input(
                    f"Row cap ({nm})",
                    min_value=200,
                    max_value=20000,
                    value=3000,
                    step=200,
                    key=cap_key
                )
                df_show = df[final_cols]
                if len(df_show) > cap:
                    st.warning(f"{nm}: showing first {cap} rows out of {len(df_show):,}.")
                    df_show = df_show.head(int(cap))

                st.dataframe(df_show, use_container_width=True, height=420)

                dl_key = make_key("dl_single", cid, nm, sheet, idx)
                csv = df_show.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    f"Download CSV ({nm} - {sheet})",
                    data=csv,
                    file_name=f"{cid}__{nm}__{sheet}.csv",
                    mime="text/csv",
                    key=dl_key
                )

    # ---- Sub-tab 3: Explore (Collections Browser) with collection → files hierarchy ----
    with sub_tab_browser:
        st.caption("Browse by hierarchy: collection_id (level 1) → files (level 2). Each file displays in its own panel (no merge).")

        selected_cids = st.multiselect(
            "Collections to display",
            options=collections_all,
            default=[cid] if cid in collections_all else collections_all[:1],
            key=make_key("ms_collections_browser")
        )
        if not selected_cids:
            st.info("Select at least one collection to display.")
            st.stop()

        for cidx, ccid in enumerate(selected_cids):
            st.markdown(f"### 📂 Collection: `{ccid}`")

            file_paths = [Path(p) for p in list_files_for_collection(ccid)]
            if not file_paths:
                st.info("No Excel files in this collection.")
                continue

            names = [p.name for p in file_paths]
            default_files = names[:min(2, len(names))]

            ms_files_key = make_key("ms_files_browser", ccid)
            chosen_files = st.multiselect(
                f"Files in {ccid}",
                options=names,
                default=default_files,
                key=ms_files_key
            )
            if not chosen_files:
                st.info(f"Select at least one file in `{ccid}` to display.")
                continue

            name2path = {p.name: str(p) for p in file_paths}

            for fidx, fn in enumerate(chosen_files):
                pstr = name2path.get(fn)
                if not pstr or not Path(pstr).exists():
                    st.warning(f"File not found: {fn}")
                    continue

                mtime = Path(pstr).stat().st_mtime
                try:
                    sheets = get_sheet_names_for_file(pstr, mtime)
                except Exception as e:
                    st.warning(f"Unable to open {fn}: {e}")
                    continue
                if not sheets:
                    st.warning(f"No readable sheets in {fn}.")
                    continue

                with st.expander(f"📄 {fn}", expanded=False):
                    sheet_key = make_key("sel_sheet_browser", ccid, fn, fidx)
                    sheet = st.selectbox(
                        f"Sheet to view ({ccid}/{fn})",
                        options=sheets,
                        index=0,
                        key=sheet_key
                    )

                    try:
                        df = load_excel_sheet(pstr, sheet, mtime).copy()
                    except Exception as e:
                        st.warning(f"Unable to read {fn}/{sheet}: {e}")
                        continue

                    if df is None or df.empty:
                        st.info("This sheet has no rows.")
                        continue

                    all_cols = list(df.columns)
                    cols_key = make_key("ms_cols_browser", ccid, fn, fidx)
                    picked = st.multiselect(
                        f"Columns to show ({ccid}/{fn})",
                        options=all_cols,
                        default=all_cols,
                        key=cols_key
                    )
                    final_cols = picked if picked else all_cols

                    cap_key = make_key("num_cap_browser", ccid, fn, fidx)
                    cap = st.number_input(
                        f"Row cap ({ccid}/{fn})",
                        min_value=200,
                        max_value=20000,
                        value=3000,
                        step=200,
                        key=cap_key
                    )
                    df_show = df[final_cols]
                    if len(df_show) > cap:
                        st.warning(f"{ccid}/{fn}: showing first {cap} rows out of {len(df_show):,}.")
                        df_show = df_show.head(int(cap))

                    st.dataframe(df_show, use_container_width=True, height=420)

                    dl_key = make_key("dl_browser", ccid, fn, sheet, fidx)
                    csv = df_show.to_csv(index=False).encode("utf-8-sig")
                    st.download_button(
                        f"Download CSV ({ccid} - {fn} - {sheet})",
                        data=csv,
                        file_name=f"{ccid}__{fn}__{sheet}.csv",
                        mime="text/csv",
                        key=dl_key
                    )
