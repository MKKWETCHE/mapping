import streamlit as st
import pandas as pd
from io import BytesIO
import os
import re
from typing import Dict

# -----------------------------
# Page configuration
# -----------------------------
st.set_page_config(
    layout="wide",
    page_title="Muuto Mapping Lookup",
    page_icon="favicon.png",
)

# -----------------------------
# Styling
# -----------------------------
st.markdown(
    """
<style>
    .stApp, body { background-color: #EFEEEB !important; }
    .main .block-container { background-color: #EFEEEB !important; padding-top: 2rem; }
    h1, h2, h3 { text-transform: none !important; }
    h1 { color: #333; }
    h2 { color: #1E40AF; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; }
    h3 { color: #1E40AF; font-size: 1.25em; padding-bottom: 3px; margin-top: 20px; margin-bottom: 10px; }

    div[data-testid="stAlert"] { background-color: #f0f2f6 !important; border: 1px solid #D1D5DB !important; border-radius: 0.25rem !important; }
    div[data-testid="stAlert"] > div:first-child { background-color: transparent !important; }

    div[data-testid="stTextArea"] textarea { background-color: #FFFFFF !important; color: #000000 !important; border: 1px solid #CCCCCC !important; }
    div[data-testid="stTextArea"] textarea:focus { border-color: #5B4A14 !important; box-shadow: 0 0 0 1px #5B4A14 !important; }

    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"] {
        border: 1px solid #5B4A14 !important; background-color: #FFFFFF !important; color: #5B4A14 !important;
        padding: 0.375rem 0.75rem !important; font-size: 1rem !important; border-radius: 0.25rem !important; font-weight: 500 !important;
    }
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:hover {
        background-color: #5B4A14 !important; color: #FFFFFF !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------------
# Constants
# -----------------------------
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "muuto_logo.png")
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQPRmVmc0LYISduQdJyfz-X3LJlxiEDCNwW53LhFsWp5fFDS8V669rCd9VGoygBZSAZXeSNZ5fquPen/pub?output=csv"
OUTPUT_HEADERS = ["New Item No.", "OLD Item-variant", "Ean no.", "Description", "Family", "Category"]

# -----------------------------
# Helpers
# -----------------------------
def parse_pasted_ids(raw: str):
    if not raw:
        return []
    tokens = re.split(r"[\s,;]+", raw.strip())
    cleaned = [t.strip().strip('"').strip("'") for t in tokens if t.strip()]
    seen, out = set(), []
    for t in cleaned:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out

@st.cache_data(show_spinner=False)
def read_mapping_from_gsheets(csv_url: str) -> pd.DataFrame:
    try:
        df = pd.read_csv(csv_url, dtype=str, keep_default_na=False)
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Failed to read Google Sheets CSV export: {e}")
        return pd.DataFrame()

def map_case_insensitive(df: pd.DataFrame, required: list) -> Dict[str, str]:
    lower_map = {c.lower(): c for c in df.columns}
    return {name: lower_map.get(name.lower()) for name in required}

def select_order_and_rename(df: pd.DataFrame, colmap: Dict[str, str]) -> pd.DataFrame:
    cols = []
    for h in OUTPUT_HEADERS:
        actual = colmap.get(h)
        if actual and actual in df.columns:
            cols.append(actual)
        else:
            df[h] = None
            cols.append(h)
    out = df[cols].copy()
    rename_map = {colmap[h]: h for h in OUTPUT_HEADERS if colmap.get(h) and colmap[h] != h}
    if rename_map:
        out = out.rename(columns=rename_map)
    return out

def to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Lookup Output") -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()

# -----------------------------
# Header
# -----------------------------
left, right = st.columns([6, 1])
with left:
    st.title("Muuto Mapping Lookup")
    st.markdown(
        """
**What this app does**

1. Paste a list of IDs (Muuto item-variant numbers or EANs).
2. The app looks up each ID in the internal mapping sheet stored online.
3. It returns these columns: `New Item No.`, `OLD Item-variant`, `Ean no.`, `Description`, `Family`, `Category`.
4. Download the result as an Excel file.
        """
    )
with right:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)

st.markdown("---")

# -----------------------------
# Inputs
# -----------------------------
st.subheader("Paste IDs")
raw_input = st.text_area(
    "Paste IDs",
    height=200,
    placeholder="Example:\n5710562801234\nMTO-CHAIR-001-01\n5710562805678\nMTO-SOFA-CHAIS-LEFT-22",
)
ids = parse_pasted_ids(raw_input)

# -----------------------------
# Load Mapping
# -----------------------------
mapping_df = read_mapping_from_gsheets(DEFAULT_SHEET_URL)
if mapping_df.empty:
    st.error("Failed to load mapping sheet. Make sure the sheet is published to the web with CSV output.")
    st.stop()

required = OUTPUT_HEADERS + ["OLD Item-variant", "Ean no."]
colmap = map_case_insensitive(mapping_df, required)
if not colmap.get("OLD Item-variant") or not colmap.get("Ean no."):
    st.error("Required columns not found: 'OLD Item-variant' and/or 'Ean no.'")
    st.stop()

old_col = colmap["OLD Item-variant"]
ean_col = colmap["Ean no."]
work = mapping_df.copy()
work[old_col] = work[old_col].astype(str).str.strip()
work[ean_col] = work[ean_col].astype(str).str.strip()

# -----------------------------
# Lookup
# -----------------------------
st.subheader("Lookup")
if not ids:
    st.info("Paste IDs to run the lookup.")
else:
    mask = work[old_col].isin(ids) | work[ean_col].isin(ids)
    matches = work.loc[mask].copy()
    matched_keys = set(matches[old_col].dropna().astype(str)) | set(matches[ean_col].dropna().astype(str))
    not_found = [x for x in ids if x not in matched_keys]

    ordered = select_order_and_rename(matches, colmap)

    c1, c2, c3 = st.columns([1, 1, 4])
    with c1:
        st.metric("IDs provided", len(ids))
    with c2:
        st.metric("Matches", len(ordered))
    with c3:
        if not_found:
            st.caption("IDs without a match:")
            st.code("\n".join(not_found), language=None)

    st.dataframe(ordered, use_container_width=True, hide_index=True)

    xlsx = to_xlsx_bytes(ordered)
    st.download_button(
        label="Download Excel",
        data=xlsx,
        file_name="muuto_mapping_lookup.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# -----------------------------
# Footer
# -----------------------------
st.markdown(
    """
<small>
The mapping file is loaded automatically from the Muuto Google Sheet (published as CSV). Leading zeros are preserved.
</small>
""",
    unsafe_allow_html=True,
)
