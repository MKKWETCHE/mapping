import streamlit as st
import pandas as pd
from io import BytesIO
import os
import re
from typing import Dict, List

# --- Fil-afhængige konstanter (sikrer stien virker, uanset hvor den køres fra) ---
# Skal erstattes med den faktiske sti til dit logo
try:
    # Denne sti er kun relevant, hvis du kører lokalt, men holder den for konsistens.
    # I et Streamlit cloud-miljø skal du sørge for, at 'muuto_logo.png' er i samme mappe.
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")
except NameError:
    # Nogle miljøer (som Jupyter/Colab) har ikke __file__
    LOGO_PATH = "muuto_logo.png"

# -----------------------------
# Konstante værdier
# -----------------------------
# ERSTAT MED DEN AKTUELLE GOOGLE SHEETS URL
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1S50it_q1BahpZCPW8dbuN7DyOMnyDgFIg76xIDSoXEk/edit?usp=sharing"

OUTPUT_HEADERS = [
    "New Item No.",
    "OLD Item-variant",
    "Ean no.",
    "Description",
    "Family",
    "Category",
]

# -----------------------------
# Sidekonfiguration
# -----------------------------
st.set_page_config(
    layout="wide",
    page_title="Muuto Varenummer Konvertering",
    page_icon="📦",  # Skift til et passende ikon, hvis favicon.png ikke er tilgængeligt
)

# -----------------------------
# Styling (tilpasset Muuto's brandfarver)
# Jeg har valgt en neutral baggrund og en dyb brun/guld for branding
# for at skabe et mere eksklusivt look.
# -----------------------------
# Baggrund: EFEEEB (Meget lys varm grå)
# Accent: 5B4A14 (Dyb, mættet brun/guld)
# Tekst: 333 (Mørkegrå)

st.markdown(
    """
<style>
    .stApp, body { background-color: #EFEEEB !important; }
    .main .block-container { background-color: #EFEEEB !important; padding-top: 2rem; }
    h1, h2, h3 { text-transform: none !important; }
    h1 { color: #5B4A14; font-size: 2.5em; margin-top: 0; }
    h2 { color: #333; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; border-bottom: 1px solid #CCC; }
    h3 { color: #5B4A14; font-size: 1.5em; padding-bottom: 3px; margin-top: 20px; margin-bottom: 10px; }
    h4 { color: #333; font-size: 1.1em; margin-top: 15px; margin-bottom: 5px; }

    /* Forbedret advarselsboks */
    div[data-testid="stAlert"] { background-color: #f7f6f4 !important; border: 1px solid #dcd4c3 !important; border-radius: 0.25rem !important; }
    div[data-testid="stAlert"] > div:first-child { background-color: transparent !important; }
    div[data-testid="stAlert"] div[data-testid="stMarkdownContainer"],
    div[data-testid="stAlert"] div[data-testid="stMarkdownContainer"] p { color: #31333F !important; }
    div[data-testid="stAlert"] svg { fill: #5B4A14 !important; }

    /* Inputs */
    div[data-testid="stTextArea"] textarea,
    div[data-testid="stTextInput"] input,
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:first-child,
    div[data-testid="stMultiSelect"] div[data-baseweb="input"],
    div[data-testid="stMultiSelect"] > div > div[data-baseweb="select"] > div:first-child {
        background-color: #FFFFFF !important; color: #000000 !important; border: 1px solid #CCCCCC !important;
    }
    div[data-testid="stTextArea"] textarea:focus,
    div[data-testid="stTextInput"] input:focus,
    div[data-testid="stSelectbox"] div[data-baseweb="select"][aria-expanded="true"] > div:first-child,
    div[data-testid="stMultiSelect"] div[data-baseweb="input"]:focus-within,
    div[data-testid="stMultiSelect"] div[aria-expanded="true"] {
        border-color: #5B4A14 !important; box-shadow: 0 0 0 1px #5B4A14 !important;
    }

    /* Buttons (Muuto Guld/Brun) */
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"],
    div[data-testid="stButton"] button[data-testid^="stBaseButton"] {
        border: 1px solid #5B4A14 !important; background-color: #5B4A14 !important; color: #FFFFFF !important;
        padding: 0.5rem 1rem !important; font-size: 1rem !important; line-height: 1.5 !important; border-radius: 0.25rem !important;
        transition: color 0.15s ease-in-out, background-color 0.15s ease-in-out, border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out !important; font-weight: 600 !important;
        text-transform: uppercase !important;
    }
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:hover,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"]:hover {
        background-color: #4A3D10 !important; color: #FFFFFF !important; border-color: #4A3D10 !important;
    }
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:active,
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:focus,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"]:active,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"]:focus {
        background-color: #4A3D10 !important; color: #FFFFFF !important; border-color: #4A3D10 !important; box-shadow: 0 0 0 0.2rem rgba(91, 74, 20, 0.4) !important; outline: none !important;
    }
    
    /* Datatabel styling (lettere at læse) */
    .stDataFrame {
        border: 1px solid #CCC;
        border-radius: 0.25rem;
    }
    
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------------
# Hjælpefunktioner
# -----------------------------
def parse_pasted_ids(raw: str) -> List[str]:
    """Uddrager unikke varenumre fra en tekstblok."""
    if not raw:
        return []
    # Opdel efter mellemrum, kommaer og semikoloner
    tokens = re.split(r"[\s,;]+", raw.strip())
    # Fjern anførselstegn og strip mellemrum
    cleaned = [t.strip().strip('"').strip("'") for t in tokens if t.strip()]
    # Returner kun unikke ID'er
    seen, out = set(), []
    for t in cleaned:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out


def to_csv_export_url(url: str) -> str:
    """Konverterer en Google Sheets-URL til en direkte CSV-eksport-URL."""
    if not url:
        return ""
    url = url.strip()
    if "export?format=csv" in url:
        return url
    m = re.search(r"https://docs.google.com/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not m:
        return url  # Returner som den er; kan allerede være et offentligt CSV-link
    sheet_id = m.group(1)
    gid_match = re.search(r"[?&#]gid=(\d+)", url)
    gid = gid_match.group(1) if gid_match else "0"
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


@st.cache_data(show_spinner="Henter og behandler mapping-data...")
def read_mapping_from_gsheets(csv_url: str) -> pd.DataFrame:
    """Indlæser mapping-data fra en Google Sheets CSV-eksport."""
    if not csv_url:
        return pd.DataFrame()
    try:
        # Læser alt som streng for at bevare førende nuller
        df = pd.read_csv(csv_url, dtype=str, keep_default_na=False)
        # Strip whitespace fra alle celler
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"❌ Fejl ved indlæsning af Google Sheets: Kontrollér URL og delingsindstillinger. Detaljer: {e}")
        return pd.DataFrame()


def map_case_insensitive(df: pd.DataFrame, required: list) -> Dict[str, str]:
    """Mapper de påkrævede header-navne (case-insensitive) til de faktiske kolonnenavne."""
    lower_map = {c.lower(): c for c in df.columns}
    return {name: lower_map.get(name.lower()) for name in required}


def select_order_and_rename(df: pd.DataFrame, colmap: Dict[str, str]) -> pd.DataFrame:
    """Vælger de ønskede kolonner, sikrer rækkefølgen og omdøber dem til standardheadere."""
    cols = []
    # Sikr, at alle output-kolonner eksisterer og vælges i den ønskede rækkefølge
    for h in OUTPUT_HEADERS:
        actual = colmap.get(h)
        if actual and actual in df.columns:
            cols.append(actual)
        else:
            # Tilføj en tom kolonne, hvis den mangler
            df[h] = None
            cols.append(h)
            
    out = df[cols].copy()
    
    # Omdøb tilbage til de kanoniske headere
    rename_map = {colmap[h]: h for h in OUTPUT_HEADERS if colmap.get(h) and colmap[h] != h}
    if rename_map:
        out = out.rename(columns=rename_map)
        
    return out


def to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Konverteringsresultat") -> bytes:
    """Konverterer DataFrame til en Excel (.xlsx) fil i hukommelsen."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()

# -----------------------------
# App Hovedindhold
# -----------------------------

# --- Header og introduktion ---
left, right = st.columns([6, 1])
with left:
    st.title("Muuto Varenummer Konvertering")
    st.markdown("---") # Brugerdefinerede streger er nu defineret i CSS
    st.markdown(
        """
        **Velkommen til dit værktøj til nem konvertering af varenumre.**
        
        Brug dette værktøj til hurtigt at **identificere de nye varenumre** baseret på dine gamle Muuto vare-varianter eller EAN-numre.
        """
    )
with right:
    if os.path.exists(LOGO_PATH):
        # Vis logoet, hvis det findes
        st.image(LOGO_PATH, width=120)

st.markdown("---")

# -----------------------------
# Trin 1: Data Opsætning
# -----------------------------
st.header("1. Data Opsætning: Hvor er dit Mapping Sheet? ⚙️")

st.info(
    "**Bemærk:** Værktøjet kræver et Google Sheet, der indeholder alle dine nye og gamle varenumre. "
    "Sheetet skal indeholde kolonnerne **'OLD Item-variant'** og **'Ean no.'** for at kunne matche."
)

gsheets_url_raw = st.text_input(
    "Indsæt Google Sheets Link",
    value=DEFAULT_SHEET_URL,
    placeholder="Indsæt et link som f.eks. https://docs.google.com/spreadsheets/d/....",
    help=(
        "Sørg for, at dit Sheet er delt som 'Alle med linket' (læser) eller er 'Udgivet til internettet' "
        "via Filer -> Del -> Udgiv til internettet. Appen konverterer automatisk linket."
    ),
)

# --- Indlæs data ---
csv_url = to_csv_export_url(gsheets_url_raw)
mapping_df = read_mapping_from_gsheets(csv_url) if csv_url else pd.DataFrame()

if mapping_df.empty:
    st.error("⚠️ Kan ikke indlæse data. Tjek venligst dit link og delingsindstillinger.")
    st.stop()

# --- Valider kolonner ---
required = OUTPUT_HEADERS + ["OLD Item-variant", "Ean no."]
colmap = map_case_insensitive(mapping_df, required)

# Validering af påkrævede lookup-kolonner
if not colmap.get("OLD Item-variant") or not colmap.get("Ean no."):
    st.error(
        f"❌ Mangler påkrævede lookup-kolonner. Sørg for at dit Sheet indeholder **'OLD Item-variant'** og **'Ean no.'**."
        f" (Case-insensitive søgning er brugt)."
    )
    st.stop()

# Forbered lookup-kolonner som strenge
old_col = colmap["OLD Item-variant"]
ean_col = colmap["Ean no."]
# Kopi til at arbejde med, for at undgå at ændre cache-data
work = mapping_df.copy()
work[old_col] = work[old_col].astype(str).str.strip()
work[ean_col] = work[ean_col].astype(str).str.strip()


# -----------------------------
# Trin 2: Indsæt Varenumre
# -----------------------------
st.header("2. Indsæt Varenumre 📝")

raw_input = st.text_area(
    "Indsæt dine gamle vare-varianter (OLD Item-variant) eller EAN-numre her.",
    height=200,
    placeholder="Indsæt et eller flere ID'er pr. linje, adskilt af mellemrum, kommaer eller nye linjer.\n"
                "Eksempel:\n"
                "5710562801234\n"
                "MTO-CHAIR-001-01\n"
                "5710562805678\n"
                "MTO-SOFA-CHAIS-LEFT-22",
)

ids = parse_pasted_ids(raw_input)

# -----------------------------
# Trin 3: Resultater og Eksport
# -----------------------------
st.header("3. Resultater og Eksport 📊")

if not ids:
    st.info("⬆️ Indsæt dine varenumre i Trin 2 for at starte konverteringen.")
else:
    # --- Lookup Logik ---
    # Find rækker, hvor enten den gamle vare-variant ELLER EAN-nummer matcher et af de indtastede ID'er
    mask = work[old_col].isin(ids) | work[ean_col].isin(ids)
    matches = work.loc[mask].copy()

    # Identificer matchede og ikke-fundne nøgler
    # Skaber et sæt af de faktiske ID'er, der blev fundet i enten OLD Item-variant eller EAN no.
    matched_keys = set(matches[old_col].dropna().astype(str)) | set(matches[ean_col].dropna().astype(str))
    # Filtrer de oprindelige ID'er, der ikke blev fundet
    not_found = [x for x in ids if x not in matched_keys]

    # Vælg og omdøb kolonner
    ordered = select_order_and_rename(matches, colmap)

    # --- Metrics og Feedback ---
    c1, c2, c3 = st.columns([1, 1, 4])
    with c1:
        st.metric("ID'er Indtastet", len(ids))
    with c2:
        st.metric("Antal Match", len(ordered))
    with c3:
        if not_found:
            st.warning(f"⚠️ **{len(not_found)} ID'er** blev ikke matchet. Se listen nedenfor.")

    # --- Visning af ikke-fundne ---
    if not_found:
        st.caption("Følgende ID'er kunne ikke findes i dit Mapping Sheet (tjek for tastefejl):")
        st.code("\n".join(not_found), language=None)
        st.markdown("---")
        
    if ordered.empty:
        st.error("Ingen af de indtastede ID'er blev matchet i dit Sheet. Tjek venligst dine indtastninger og Sheet-data.")
        st.stop()


    # --- Resultattabel og Download ---
    st.subheader("Konverteringsresultat")
    st.dataframe(
        ordered, 
        use_container_width=True, 
        hide_index=True,
        # Gør det muligt for kunden at kopiere hele tabellen
        
    )

    xlsx = to_xlsx_bytes(ordered)
    st.download_button(
        label="Download Resultat som Excel-fil (.xlsx)",
        data=xlsx,
        file_name="muuto_varenummer_konvertering.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_button" # Streamlit-nøgle for at undgå fejl
    )

# --- Footnote ---
st.markdown("---")
st.markdown(
    """
<div style="text-align: center;">
<small>
Dette værktøj er leveret af Muuto for at lette overgangen til nye varenumre. Spørgsmål? Kontakt din salgsrepræsentant.
</small>
</div>
""",
    unsafe_allow_html=True,
)
