import streamlit as st
import pandas as pd
from io import BytesIO
import os
import re
from typing import List
import zipfile
from collections import defaultdict

# --- Paths og konstanter ---
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")
except NameError:
    BASE_DIR = "."
    LOGO_PATH = "muuto_logo.png"

MAPPING_ZIP_PATH = os.path.join(BASE_DIR, "mapping.csv.zip")
MAPPING_FILENAME = "mapping.csv"

# Kun de kolonner, du vil have i output (Description fjernet)
OUTPUT_HEADERS = [
    "New Item No.",
    "Old Item no.",
    "Ean No.",
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
        h2 { color: #333 !important; border-bottom: 1px solid #CCC; }

        div[data-testid="stDownloadButton"] button,
        div[data-testid="stButton"] button {
            border: 1px solid #5B4A14 !important;
            background-color: #5B4A14 !important;
            padding: 0.5rem 1rem !important;
            font-size: 1rem !important;
            border-radius: 0.25rem !important;
            font-weight: 600 !important;
            text-transform: uppercase !important;
            color: #FFFFFF !important;
        }

        div[data-testid="stDownloadButton"] p {
            color: #FFFFFF !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# Hjælpefunktioner
# ---------------------------------------------------------
def parse_pasted_ids(raw: str) -> List[str]:
    if not raw:
        return []
    tokens = re.split(r"[\s,;]+", raw.strip())
    seen = set()
    out = []
    for t in tokens:
        t = t.strip().strip('"').strip("'")
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def autodetect_separator(first_chunk: str) -> str:
    if ";" in first_chunk:
        return ";"
    if "\t" in first_chunk:
        return "\t"
    if "," in first_chunk:
        return ","
    return ";"


def normalize_id(s: str) -> str:
    s = str(s).strip()
    if not s:
        return ""

    if re.fullmatch(r"\d+", s):
        no_leading = s.lstrip("0")
        return no_leading or "0"

    return s


@st.cache_data(show_spinner=False)
def read_mapping_from_zip(zip_path: str, filename: str) -> pd.DataFrame:
    if not os.path.exists(zip_path):
        st.error("mapping.csv.zip not found in repository.")
        return pd.DataFrame()

    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            if filename not in zf.namelist():
                st.error(f"ZIP file does not contain {filename}")
                return pd.DataFrame()

            with zf.open(filename) as f:
                head = f.read(5000).decode("utf-8", errors="ignore")
                sep = autodetect_separator(head)

            with zf.open(filename) as f:
                df = pd.read_csv(
                    f,
                    dtype=str,
                    encoding="utf-8",
                    sep=sep,
                    engine="python",
                )

        df.columns = [c.strip() for c in df.columns]
        for c in df.columns:
            df[c] = df[c].astype(str).str.strip()

        return df

    except Exception as e:
        st.error(f"Failed to read mapping.csv.zip: {e}")
        return pd.DataFrame()


def to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def build_index(df: pd.DataFrame) -> dict:
    index_map = defaultdict(list)

    for i, val in df[OLD_COL_NAME].items():
        key = normalize_id(val)
        if key:
            index_map[key].append(i)

    for i, val in df[EAN_COL_NAME].items():
        key = normalize_id(val)
        if key:
            index_map[key].append(i)

    return index_map


def exact_lookup(ids: List[str], df: pd.DataFrame) -> pd.DataFrame:
    index_map = build_index(df)
    rows = []

    for raw_id in ids:
        key = normalize_id(raw_id)
        idxs = index_map.get(key, [])

        if idxs:
            tmp = df.loc[idxs].copy()
            tmp["Query"] = raw_id
            tmp["Match Type"] = "Exact"
            rows.append(tmp)
        else:
            empty_row = {h: None for h in OUTPUT_HEADERS}
            empty_row["Query"] = raw_id
            empty_row["Match Type"] = "No match"
            rows.append(pd.DataFrame([empty_row]))

    result = pd.concat(rows, ignore_index=True)

    for h in OUTPUT_HEADERS:
        if h not in result.columns:
            result[h] = None

    return result


# ---------------------------------------------------------
# UI
# ---------------------------------------------------------
left, right = st.columns([6, 1])
with left:
    st.title("Muuto Item Number Converter")
    st.markdown(
        """
        **Welcome to the Muuto Item Number Converter**

        This tool helps convert old Muuto item numbers to the new structure.
        Your **EAN numbers remain unchanged**, while old item numbers are converted.

        **How to use the tool:**
        1. Copy the Muuto item numbers or EAN codes from your system.  
        2. Paste them below (one per line or separated by space/comma/semicolon).  
        3. The tool will return:
            - The new Muuto item number  
            - The corresponding EAN  

        If an item number cannot be matched, it may be discontinued or contain an error.  
        Please contact **customercare@muuto.com** if you need assistance.
        """
    )

with right:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)

st.header("Copy/Paste Item Numbers or EAN Codes")

raw_input = st.text_area(
    "Paste Old Muuto Item Numbers or EAN codes here:",
    height=200,
    key="ids_input",
)

ids = parse_pasted_ids(raw_input)
submitted = st.button("Convert IDs")

if submitted:
    if not ids:
        st.error("You must paste at least one ID before converting.")
    else:
        with st.spinner("Converting IDs..."):
            mapping_df = read_mapping_from_zip(MAPPING_ZIP_PATH, MAPPING_FILENAME)

            if mapping_df.empty:
                st.error("Mapping file is empty or unreadable.")
            else:
                missing = [c for c in [OLD_COL_NAME, EAN_COL_NAME] if c not in mapping_df.columns]
                if missing:
                    st.error(
                        f"Required column(s) missing in mapping.csv: {missing}\n"
                        f"Actual columns: {list(mapping_df.columns)}"
                    )
                else:
                    for h in OUTPUT_HEADERS:
                        if h not in mapping_df.columns:
                            mapping_df[h] = None

                    results = exact_lookup(ids, mapping_df)
                    matches_count = int((results["Match Type"] != "No match").sum())

                    results_sorted = results.sort_values(
                        by=["Match Type", "Query"],
                        ascending=[True, True],
                    )

                    results_sorted = results_sorted.rename(columns={"Query": "Your Input"})
                    display_cols = ["Your Input", "Match Type"] + OUTPUT_HEADERS
                    display_df = results_sorted[display_cols]

                    st.session_state["results_df"] = display_df
                    st.session_state["matches_count"] = matches_count
                    st.session_state["ids_count"] = len(ids)

if "results_df" in st.session_state:
    display_df = st.session_state["results_df"]
    matches_count = st.session_state.get("matches_count", 0)
    ids_count = st.session_state.get("ids_count", 0)

    st.header("2. Results")
    st.metric("IDs Provided", ids_count)
    st.metric("IDs with a match", matches_count)

    st.dataframe(display_df, use_container_width=True, hide_index=True)

    xlsx = to_xlsx_bytes(display_df)
    st.download_button(
        "Download Excel File",
        data=xlsx,
        file_name="muuto_item_conversion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Paste your IDs above and click **Convert IDs** to run the lookup.")
