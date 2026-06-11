"""
CleanDRR — Streamlit web app version.

Deploy on Streamlit Cloud:
  1. Push this file + requirements.txt to a GitHub repo
  2. requirements.txt should contain:
        streamlit
        pandas
        openpyxl
  3. Set "Main file path" to CleanDRR.py

Run locally:
    pip install streamlit pandas openpyxl
    streamlit run CleanDRR.py
"""

import io
import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill, Font

# ── CONFIG ────────────────────────────────────────────────────────────────────
STATUS_COL_NAME    = "Status"
STATUS_COL_INDEX   = 9
DATE_COL_NAME      = "Date"
TIME_COL_NAME      = "Time"
COL_E_INDEX        = 4          # Account Number column
DIALED_COL_INDEX   = 5          # Column F = Dialed Number
PTP_AMOUNT_COL     = "PTP Amount"
CLAIM_PAID_COL     = "Claim Paid Amount"
REMARK_COL         = "Remark"
REMOVE_STATUSES    = {"BP", "REACTIVE", "SMS FAILED", "NEW", "ABORTED"}
COLS_TO_DROP_START = 27
COLS_TO_DROP_END   = 50


# ── PROCESSING LOGIC (unchanged from desktop version) ────────────────────────
def process_file(file_like):
    df_headers = pd.read_excel(file_like, nrows=0)

    str_cols = {}
    if COL_E_INDEX < len(df_headers.columns):
        str_cols[df_headers.columns[COL_E_INDEX]] = str

    col_e_name      = df_headers.columns[COL_E_INDEX]      if COL_E_INDEX      < len(df_headers.columns) else None
    dialed_col_name = df_headers.columns[DIALED_COL_INDEX] if DIALED_COL_INDEX < len(df_headers.columns) else None

    if dialed_col_name:
        str_cols[dialed_col_name] = str

    for col in df_headers.columns:
        if "dialed" in str(col).lower():
            str_cols[col] = str
            dialed_col_name = col
            break

    # Re-seek the buffer since we read headers above
    if hasattr(file_like, "seek"):
        file_like.seek(0)
    df = pd.read_excel(file_like, dtype=str_cols)

    def strip_decimal(x):
        if pd.isna(x) or str(x).strip() in ("", "nan"):
            return ""
        s = str(x).strip()
        if "." in s:
            integer_part = s.split(".")[0]
            if integer_part.lstrip("-").isdigit():
                return integer_part
        return s

    def clean_account(x):
        if pd.isna(x) or str(x).strip() in ("", "nan"):
            return ""
        s = str(x).strip()
        if s.endswith(".0") and s[:-2].lstrip("-").isdigit() and not s[:-2].startswith("0"):
            return s[:-2]
        return s

    if col_e_name and col_e_name in df.columns:
        df[col_e_name] = df[col_e_name].apply(clean_account)

    if dialed_col_name and dialed_col_name in df.columns:
        df[dialed_col_name] = df[dialed_col_name].apply(strip_decimal)
    else:
        for col in df.columns:
            if "dialed" in str(col).lower():
                df[col] = df[col].apply(strip_decimal)
                dialed_col_name = col
                break

    STATUS_COL = STATUS_COL_NAME if STATUS_COL_NAME in df.columns else (
        df.columns[STATUS_COL_INDEX] if STATUS_COL_INDEX < len(df.columns) else None
    )
    if not STATUS_COL:
        raise ValueError("Status column not found.")

    status_normalized = df[STATUS_COL].astype(str).str.strip().str.upper()
    is_blank          = df[STATUS_COL].isna() | (df[STATUS_COL].astype(str).str.strip() == "")
    has_cease         = status_normalized.str.contains("CEASE", na=False)
    is_removable      = status_normalized.isin(REMOVE_STATUSES) | is_blank | has_cease

    if PTP_AMOUNT_COL in df.columns:
        has_ptp       = status_normalized.str.contains(r"\bPTP\b", regex=True, na=False)
        ptp_numeric   = pd.to_numeric(df[PTP_AMOUNT_COL].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
        ptp_has_value = has_ptp & (ptp_numeric > 0)
        ptp_no_value  = has_ptp & (ptp_numeric <= 0)
        is_removable  = (is_removable | ptp_no_value) & ~ptp_has_value

    if CLAIM_PAID_COL in df.columns:
        has_kept       = status_normalized.str.contains(r"\bKEPT\b", regex=True, na=False)
        claim_numeric  = pd.to_numeric(df[CLAIM_PAID_COL].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
        kept_has_value = has_kept & (claim_numeric > 0)
        kept_no_value  = has_kept & (claim_numeric <= 0)
        is_removable   = (is_removable | kept_no_value) & ~kept_has_value

    cleaned_df = df[~is_removable].copy()
    removed_df = df[is_removable].copy()

    def get_reason(row):
        s = str(row[STATUS_COL]).strip().upper()
        if s in ("", "NAN"):
            return "Blank Status"
        if s in REMOVE_STATUSES:
            return f"Status: {str(row[STATUS_COL]).strip()}"
        if "CEASE" in s:
            return f"Status contains CEASE: {str(row[STATUS_COL]).strip()}"
        if "PTP" in s:
            return "PTP with no PTP Amount"
        if "KEPT" in s:
            return "KEPT with no Claim Paid Amount"
        return f"Status: {str(row[STATUS_COL]).strip()}"

    removed_df.insert(0, "Removed Reason", removed_df.apply(get_reason, axis=1))

    srp_mask   = pd.Series([False] * len(cleaned_df), index=cleaned_df.index)
    remarks_df = pd.DataFrame()

    if REMARK_COL in cleaned_df.columns:
        remark_norm    = cleaned_df[REMARK_COL].astype(str).str.strip().str.upper()
        status_clean   = cleaned_df[STATUS_COL].astype(str).str.strip().str.upper()
        has_action_ptp = remark_norm.str.contains(r"ACTION\s*:\s*PTP", regex=True, na=False)
        status_not_ptp_kept = (
            ~status_clean.str.contains("PTP",  na=False) &
            ~status_clean.str.contains("KEPT", na=False)
        )
        srp_mask = has_action_ptp & status_not_ptp_kept

        if srp_mask.any():
            changed_rows = cleaned_df.loc[srp_mask].copy()
            remarks_df = pd.DataFrame({
                "Row #":        range(1, srp_mask.sum() + 1),
                STATUS_COL:     changed_rows[STATUS_COL].values,
                REMARK_COL + " (Before)": changed_rows[REMARK_COL].astype(str).values,
                REMARK_COL + " (After)":  changed_rows[REMARK_COL].astype(str).str.replace(
                    r"(?i)Action\s*:\s*PTP", "Action: SRP", regex=True).values,
            })

        cleaned_df.loc[srp_mask, REMARK_COL] = cleaned_df.loc[srp_mask, REMARK_COL].astype(str).str.replace(
            r"(?i)Action\s*:\s*PTP", "Action: SRP", regex=True)

    cols_to_drop = [df.columns[i] for i in range(COLS_TO_DROP_START, min(COLS_TO_DROP_END + 1, len(df.columns)))]
    cleaned_df.drop(columns=cols_to_drop, inplace=True, errors="ignore")
    removed_df.drop(columns=cols_to_drop, inplace=True, errors="ignore")

    if DATE_COL_NAME in cleaned_df.columns:
        cleaned_df[DATE_COL_NAME] = pd.to_datetime(cleaned_df[DATE_COL_NAME], errors="coerce").dt.strftime("%m-%d-%Y")

    if DATE_COL_NAME in cleaned_df.columns and TIME_COL_NAME in cleaned_df.columns:
        date_str = pd.to_datetime(cleaned_df[DATE_COL_NAME], errors="coerce").dt.strftime("%m-%d-%Y")
        time_str = pd.to_datetime(cleaned_df[TIME_COL_NAME], errors="coerce").dt.strftime("%I:%M:%S %p")
        cleaned_df[TIME_COL_NAME] = date_str + " " + time_str

    stats = {
        "total":       len(df),
        "retained":    len(cleaned_df),
        "removed":     len(removed_df),
        "srp_changed": int(srp_mask.sum()),
    }

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, frame in [("Cleaned", cleaned_df), ("Removed", removed_df)]:
            frame.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                if cell.value == "Removed Reason":
                    cell.fill = PatternFill("solid", fgColor="C05621")
                else:
                    cell.fill = PatternFill("solid", fgColor="1E2235")
            if "Removed Reason" in frame.columns:
                reason_idx = list(frame.columns).index("Removed Reason") + 1
                for row in ws.iter_rows(min_row=2, min_col=reason_idx, max_col=reason_idx):
                    for cell in row:
                        cell.fill = PatternFill("solid", fgColor="FFF3E0")
                        cell.font = Font(bold=True, color="C05621")
            for col_name in [col_e_name, dialed_col_name]:
                if col_name and col_name in frame.columns:
                    col_idx = list(frame.columns).index(col_name) + 1
                    for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                        for cell in row:
                            cell.number_format = "@"
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

        if not remarks_df.empty:
            remarks_df.to_excel(writer, index=False, sheet_name="Remarks Changes")
            ws = writer.sheets["Remarks Changes"]
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="7C3AED")
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    buf.seek(0)
    return cleaned_df, removed_df, remarks_df, stats, buf.read()


# ── STREAMLIT UI ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel Data Cleaner", page_icon="⬡", layout="wide")

st.markdown(
    """
    <style>
      .block-container { padding-top: 2rem; }
      h1 { font-weight: 700; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("⬡ Excel Data Cleaner")
st.caption("Upload · Preview · Verify · Download")

with st.sidebar:
    st.header("1. Upload file")
    uploaded = st.file_uploader(
        "Drop Excel file here",
        type=["xlsx", "xlsm", "xls"],
        help="Supported: .xlsx, .xlsm, .xls",
    )
    run = st.button("▶ Run Cleaner", type="primary", disabled=uploaded is None, use_container_width=True)

if uploaded is None:
    st.info("👈 Upload an Excel file in the sidebar to begin.")
    st.stop()

if not run and "result" not in st.session_state:
    st.success(f"Loaded **{uploaded.name}** — click **Run Cleaner** in the sidebar.")
    st.stop()

if run:
    with st.spinner("Processing file..."):
        try:
            cleaned_df, removed_df, remarks_df, stats, output_bytes = process_file(uploaded)
            st.session_state.result = {
                "cleaned": cleaned_df,
                "removed": removed_df,
                "remarks": remarks_df,
                "stats":   stats,
                "bytes":   output_bytes,
                "name":    uploaded.name,
            }
        except Exception as e:
            st.error(f"Error: {e}")
            st.stop()

result     = st.session_state.result
stats      = result["stats"]
cleaned_df = result["cleaned"]
removed_df = result["removed"]
remarks_df = result["remarks"]

# ── Stats ─────────────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
c1.metric("📋 Total Rows",    stats["total"])
c2.metric("✅ Rows Retained", stats["retained"])
c3.metric("🗑 Rows Removed",  stats["removed"])
c4.metric("✏️ Remarks Fixed", stats["srp_changed"])

# ── Download ──────────────────────────────────────────────────────────────────
out_name = result["name"].rsplit(".", 1)[0] + "_cleaned.xlsx"
st.download_button(
    "⬇  Download Cleaned Excel",
    data=result["bytes"],
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

st.divider()

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs([
    f"✅ Cleaned ({len(cleaned_df)})",
    f"🗑 Removed ({len(removed_df)})",
    f"✏️ Remarks Changes ({len(remarks_df)})",
])

def filtered(df, query):
    if not query:
        return df
    mask = df.astype(str).apply(lambda col: col.str.contains(query, case=False, na=False)).any(axis=1)
    return df[mask]

with tab1:
    q = st.text_input("🔍 Filter rows", key="q_cleaned", placeholder="Type to filter...")
    st.dataframe(filtered(cleaned_df, q), use_container_width=True, height=500)

with tab2:
    q = st.text_input("🔍 Filter rows", key="q_removed", placeholder="Type to filter...")
    st.dataframe(filtered(removed_df, q), use_container_width=True, height=500)

with tab3:
    if remarks_df.empty:
        st.info("No remarks were changed.")
    else:
        q = st.text_input("🔍 Filter rows", key="q_remarks", placeholder="Type to filter...")
        st.dataframe(filtered(remarks_df, q), use_container_width=True, height=500)
