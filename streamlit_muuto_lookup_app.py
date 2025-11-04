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
# Styling (reused)
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
    h4 { color: #102A63; font-size: 1.1em; margin-top: 15px; margin-bottom: 5px; }

    div[data-testid=\"stAlert\"] { background-color: #f0f2f6 !important; border: 1px solid #D1D5DB !important; border-radius: 0.25rem !important; }
    div[data-testid=\"stAlert\"] > div:first-child { background-color: transparent !important; }
    div[data-testid=\"stAlert\"] div[data-testid=\"stMarkdownContainer\"], div[data-testid=\"stAlert\"] div[data-testid=\"stMarkdownContainer\"] p { color: #31333F !important; }
    div[data-testid=\"stAlert\"] svg { fill: #4B5563 !important; }

    /* Inputs */
    div[data-testid=\"stTextArea\"] textarea,
    div[data-testid=\"stTextInput\"] input,
    div[data-testid=\"stSelectbox\"] div[data-baseweb=\"select\"] > div:first-child,
    div[data-testid=\"stMultiSelect\"] div[data-baseweb=\"input\"],
    div[data-testid=\"stMultiSelect\"] > div > div[data-baseweb=\"select\"] > div:first-child {
        background-color: #FFFFFF !important; color: #000000 !important; border: 1px solid #CCCCCC !important;
    }
    div[data-testid=\"stTextArea\"] textarea:focus,
    div[data-testid=\"stTextInput\"] input:focus,
    div[data-testid=\"stSelectbox\"] div[data-baseweb=\"select\"][aria-expanded=\"true\"] > div:first-child,
    div[data-testid=\"stMultiSelect\"] div[data-baseweb=\"input\"]:focus-within,
    div[data-testid=\"stMultiSelect\"] div[aria-expanded=\"true\"] {
        border-color: #5B4A14 !important; box-shadow: 0 0 0 1px #5B4A14 !important;
    }

    /* Buttons */
    div[data-testid=\"stDownloadButton\"] button[data-testid^=\"stBaseButton\"],
    div[data-testid=\"stButton\"] button[data-testid^=\"stBaseButton\"] {
        border: 1px solid #5B4A14 !important; background-color: #FFFFFF !important; color: #5B4A14 !important;
        padding: 0.375rem 0.75rem !important; font-size: 1rem !important; line-height: 1.5 !important; border-radius: 0.25rem !important;
        transition: color 0.15s ease-in-out, background-color 0.15s ease-in-out, border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out !important; font-weight: 500 !important;
        text-transform: none !important;
    }
    div[data-testid=\"stDownloadButton\"] button[data-testid^=\"stBaseButton\"]:hover,
    div[data-testid=\"stButton\"] button[data-testid^=\"stBaseButton\"]:hover { background-color: #5B4A14 !important; color: #FFFFFF !important; border-color: #5B4A14 !important; }
    div[data-testid=\"stDownloadButton\"] button[data-testid^=\"stBaseButton\"]:active,
    div[data-testid=\"stDownloadButton\"] button[data-testid^=\"stBaseButton\"]:focus,
    div[data-testid=\"stButton\"] button[data-testid^=\"stBaseButton\"]:active,
    div[data-testid=\"stButton\"] button[data-testid^=\"stBaseButton\"]:focus { background-color: #4A3D10 !important; color: #FFFFFF !important; border-color: #4A3D10 !important; box-shadow: 0 0 0 0.2rem rgba(91, 74, 20, 0.4) !important; outline: none !important; }
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------------
# Constants
# -----------------------------
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "muuto_logo.png")
OUTPUT_HEADERS = [
    "New Item No.",
    "OLD Item-variant",
    "Ean no.",
    "Description",
    "Family",
    "Category",
]

# -----------------------------
# Helpers
# -----------------------------

def parse_pasted_ids(raw: str):
    if not raw:
        return []
    tokens = re.split(r"[\s,;]+", raw.strip())
    cleaned = [t.strip().strip('\"').strip("'") for t in tokens if t.strip()]
    seen, out = set(), []
    for t in cleaned:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out


def to_csv_export_url(url: str) -> str:
    """Accept a Google Sheets URL and return a direct CSV export URL (keeps gid)."""
    if not url:
        return ""
    url = url.strip()
    if "export?format=csv" in url:
        return url
    m = re.search(r"https://docs.google.com/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not m:
        return url
    sheet_id = m.group(1)
    gid_match = re.search(r"[?&#]gid=(\d+)", url)
    gid = gid_match.group(1) if gid_match else "0"
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


@st.cache_data(show_spinner=False)
def read_mapping_from_gsheets(csv_url: str) -> pd.DataFrame:
    if not csv_url:
        return pd.DataFrame()
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
2. Provide a Google Sheets link to the mapping table.
3. The app matches each ID against **either** `OLD Item-variant` **or** `Ean no.` in the sheet.
4. It returns a table with these columns, in order: `New Item No.`, `OLD Item-variant`, `Ean no.`, `Description`, `Family`, `Category`.
5. Download the result as an Excel file.
        """
    )
with right:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)

st.markdown("---")

# -----------------------------
# Inputs
# -----------------------------
st.subheader("Inputs")

gsheets_url_raw = st.text_input(
    "Google Sheets link",
    value="https://docs.google.com/spreadsheets/d/1S50it_q1BahpZCPW8dbuN7DyOMnyDgFIg76xIDSoXEk/edit?usp=sharing",
    placeholder="Paste a link like https://docs.google.com/spreadsheets/d/....",
    help=(
        "Share your sheet as 'Anyone with the link' (Viewer) or use File → Share → Publish to the web, then paste the link here. "
        "The app converts it to a direct CSV export link automatically."
    ),
) or use File → Share → Publish to the web, then paste the link here. "
        "The app converts it to a direct CSV export link automatically."
    ),
)

raw_input = st.text_area(
    "Paste IDs",
    height=200,
    placeholder="Example:\n5710562801234\nMTO-CHAIR-001-01\n5710562805678\nMTO-SOFA-CHAIS-LEFT-22",
)

ids = parse_pasted_ids(raw_input)

# Resolve Google Sheets CSV export URL and load mapping
csv_url = to_csv_export_url(gsheets_url_raw)
mapping_df = read_mapping_from_gsheets(csv_url) if csv_url else pd.DataFrame()

if mapping_df.empty:
    st.info("Provide a valid Google Sheets link to continue.")
    st.stop()

required = OUTPUT_HEADERS + ["OLD Item-variant", "Ean no."]
colmap = map_case_insensitive(mapping_df, required)

if not colmap.get("OLD Item-variant") or not colmap.get("Ean no."):
    st.error("Required columns not found (case-insensitive): 'OLD Item-variant' and/or 'Ean no.' in your sheet.")
    st.stop()

# Prepare lookup columns as strings
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

# Footnote
st.markdown(
    """
<small>
Tip: Publishing the Google Sheet to the web guarantees a stable CSV export link. Leading zeros are preserved by reading everything as text.
</small>
""",
    unsafe_allow_html=True,
)
