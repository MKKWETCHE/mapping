import streamlit as st
import pandas as pd
from io import BytesIO
import os
import re
from typing import List
import time
import zipfile

# --- File-dependent Constants ---
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")
except NameError:
    BASE_DIR = "."
    LOGO_PATH = "muuto_logo.png"

MAPPING_CSV_ZIP = os.path.join(BASE_DIR, "mapping.csv.zip")

OUTPUT_HEADERS = [
    "New Item No.",
    "Old Item no.",
    "Ean No.",
    "Description",
    "Family",
    "Category",
]

OLD_COL_NAME = "Old Item no."
EAN_COL_NAME = "Ean No."

# ---------------------------------------------------------
# Page configuration
# ---------------------------------------------------------
st.set_page_config(
    layout="wide",
    page_title="Muuto Item Number Converter",
    page_icon="favicon.png",
)

# ---------------------------------------------------------
# CSS
# ---------------------------------------------------------
st.markdown(
    """
    <style>
        .stApp, body { background-color: #EFEEEB !important; }
        .main .block-container { background-color: #EFEEEB !important; padding-top: 2rem; }

        h1 { color: #5B4A14; font-size: 2.5em; margin-top: 0; }
        h2 { color: #333 !important; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; border-bottom: 1px solid #CCC; }
        h3 { color: #5B4A14; font-size: 1.5em; padding-bottom: 3px; margin-top: 20px; margin-bottom: 10px; }
        h4 { color: #333 !important; font-size: 1.1em; margin-top: 15px; margin-bottom: 5px; }

        div[data-testid="stDownloadButton"] p { color: white !important; }
        div[data-testid="stDownloadButton"] button,
        div[data-testid="stButton"] button {
            border: 1px solid #5B4A14 !important;
            background-color: #5B4A14 !important;
            padding: 0.5rem 1rem !important;
            font-size: 1rem !important;
            border-radius: 0.25rem !important;
            font-weight: 600 !important;
            text-transform: uppercase !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def parse_pasted_ids(raw: str) -> List[str]:
    if not raw:
        return []
    tokens = re.split(r"[\s,;]+", raw.strip())
    cleaned = [t.strip().strip('"').strip("'") for t in tokens if t.strip()]
    seen = set()
    out = []
    for t in cleaned:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out


@st.cache_data(show_spinner=False)
def read_mapping_from_csv_zip(zip_path: str) -> pd.DataFrame:
    """Load mapping.csv from mapping.csv.zip with semicolon delimiter."""
    if os.path.exists(zip_path):
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                if "mapping.csv" not in zf.namelist():
                    st.error("ZIP file does not contain mapping.csv")
                    return pd.DataFrame()

                with zf.open("mapping.csv") as f:
                    df = pd.read_csv(
                        f,
                        dtype=str,
                        encoding="utf-8",
                        sep=";",                      # <-- IMPORTANT
                        engine="python"               # <-- tolerant reader
                    )
                    df.columns = [c.strip() for c in df.columns]
                    for c in df.columns:
                        df[c] = df[c].astype(str).str.strip()
                    return df

        except Exception as e:
            st.error(f"Failed to read mapping.csv.zip: {e}")
            return pd.DataFrame()

    st.error("mapping.csv.zip not found in repository.")
    return pd.DataFrame()


def to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


# ---------------------------------------------------------
# UI
# ---------------------------------------------------------
left, right = st.columns([6, 1])
with left:
    st.title("Muuto Item Number Converter")
    st.markdown("---")
    st.markdown(
        """
        **This tool maps your legacy Item-Variants or EANs to the new Muuto item numbers.**

        **How It Works:**
        1. Paste IDs below  
        2. Click **Convert IDs**  
        3. View and download results  
        """
    )
with right:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)

st.markdown("---")

st.header("1. Paste Item IDs")
raw_input = st.text_area(
    "Paste Old Item Numbers or EANs here:",
    height=200
)

ids = parse_pasted_ids(raw_input)

# SUBMIT BUTTON
submitted = st.button("Convert IDs")

# ---------------------------------------------------------
# EXECUTE ONLY WHEN USER CLICKS
# ---------------------------------------------------------
if submitted:
    if not ids:
        st.error("You must paste at least one ID before converting.")
        st.stop()

    with st.spinner("Loading mapping file..."):
        mapping_df = read_mapping_from_csv_zip(MAPPING_CSV_ZIP)

    if mapping_df.empty:
        st.error("Mapping file is empty or unreadable.")
        st.stop()

    # Required columns
    for col in [OLD_COL_NAME, EAN_COL_NAME]:
        if col not in mapping_df.columns:
            st.error(f"Required column '{col}' missing in mapping.csv")
            st.stop()

    # Lookup
    mapping_df[OLD_COL_NAME] = mapping_df[OLD_COL_NAME].astype(str)
    mapping_df[EAN_COL_NAME] = mapping_df[EAN_COL_NAME].astype(str)

    mask = mapping_df[OLD_COL_NAME].isin(ids) | mapping_df[EAN_COL_NAME].isin(ids)
    matches = mapping_df.loc[mask].copy()

    # Reorder
    ordered = pd.DataFrame()
    for h in OUTPUT_HEADERS:
        ordered[h] = matches[h] if h in matches else None

    # Metrics
    st.header("2. Results")
    st.metric("IDs Provided", len(ids))
    st.metric("Matches Found", len(ordered))

    st.dataframe(ordered, use_container_width=True, hide_index=True)

    # Download
    xlsx = to_xlsx_bytes(ordered)
    st.download_button(
        "Download Excel File",
        data=xlsx,
        file_name="muuto_item_conversion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

