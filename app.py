import re
from collections import Counter
from io import BytesIO

import pandas as pd
import streamlit as st

# -------------------------------------------------
# Page config
# -------------------------------------------------
st.set_page_config(
    page_title="Breadcrumb Acronym Finder",
    page_icon="✅",
    layout="wide",
)

# -------------------------------------------------
# Header (clean + compact)
# -------------------------------------------------
st.markdown("## Breadcrumb Acronym Finder")
st.caption(
    "Upload an Excel breadcrumb file and scan breadcrumb columns **Level 1–11 (B–L)**. "
    "Returns **Acronym**, **Cell address**, **Cell value**, and **Breadcrumb**."
)
st.divider()

# -------------------------------------------------
# CONFIG: Expected breadcrumb columns (B–L)
# -------------------------------------------------
LEVEL_COLUMNS = [f"Level {i}" for i in range(1, 12)]
EXCEL_COLUMN_MAP = {
    "Level 1": "B", "Level 2": "C", "Level 3": "D", "Level 4": "E",
    "Level 5": "F", "Level 6": "G", "Level 7": "H", "Level 8": "I",
    "Level 9": "J", "Level 10": "K", "Level 11": "L",
}

# -------------------------------------------------
# Column normalization (CRITICAL)
# -------------------------------------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    normalized = []
    for c in df.columns:
        s = str(c)
        # Replace non-breaking / weird unicode spaces
        s = s.replace("\u00A0", " ").replace("\u2007", " ").replace("\u202F", " ")
        # Trim + collapse spaces
        s = s.strip()
        s = re.sub(r"\s+", " ", s)
        # Normalize Level headers: Level8 / Level  8 / LEVEL8 → Level 8
        s = re.sub(r"^(Level)\s*(\d+)$", r"Level \2", s, flags=re.IGNORECASE)
        normalized.append(s)

    df = df.copy()
    df.columns = normalized
    return df

# -------------------------------------------------
# Acronym rules (your final rules)
# -------------------------------------------------
SUBSCRIPT_DIGITS = set("₀₁₂₃₄₅₆₇₈₉")
SUPERSCRIPT_DIGITS = set("⁰¹²³⁴⁵⁶⁷⁸⁹¹²³")

UNIT_PREFIXES = {"mm", "cm", "km", "nm", "pm", "um", "µm", "μm"}

CANDIDATE_TOKEN_REGEX = re.compile(
    r"[A-Za-z0-9₀₁₂₃₄₅₆₇₈₉⁰¹²³⁴⁵⁶⁷⁸⁹¹²³]+(?:[-_/.][A-Za-z0-9₀₁₂₃₄₅₆₇₈₉⁰¹²³⁴⁵⁶⁷⁸⁹¹²³]+)*"
)


def is_camel_case(token: str) -> bool:
    # Exclude normal TitleCase words like Construction
    if len(token) >= 2 and token[0].isupper() and token[1:].islower():
        return False
    return any(c.islower() for c in token) and any(c.isupper() for c in token)


def looks_like_unit_prefix_camel(token: str) -> bool:
    for p in UNIT_PREFIXES:
        if token.lower().startswith(p.lower()):
            rest = token[len(p):]
            if any(c.isupper() for c in rest):
                return True
    return False


def is_acronym(token: str) -> bool:
    """
    A word is considered an acronym if:
    - It contains 2+ uppercase letters anywhere (even with lowercase/special chars)
    - OR it includes at least one uppercase letter AND one digit (3D, 5G, G4S)
    - OR it includes subscript/superscript + letters (H₂O, m²)
    - OR it matches unit-prefix CamelCase (mmWave, µmSize)
    - OR it is CamelCase like iPhone, iPad, eSIM
    """
    if not token or len(token) < 2:
        return False

    uppercase_count = sum(1 for c in token if "A" <= c <= "Z")
    has_upper = uppercase_count > 0
    has_digit = any(c.isdigit() for c in token)
    has_sub_sup = any(c in SUBSCRIPT_DIGITS or c in SUPERSCRIPT_DIGITS for c in token)
    has_letter = any(c.isalpha() for c in token)

    # 1) 2+ uppercase letters anywhere
    if uppercase_count >= 2:
        return True

    # 2) Uppercase + digit
    if has_upper and has_digit:
        return True

    # 3) Subscript/superscript + letter
    if has_sub_sup and has_letter:
        return True

    # 4) Unit prefix CamelCase (mmWave)
    if looks_like_unit_prefix_camel(token):
        return True

    # 5) CamelCase (iPhone, eSIM)
    if is_camel_case(token):
        return True

    return False


def extract_acronyms(text: str):
    if not isinstance(text, str):
        return []
    tokens = CANDIDATE_TOKEN_REGEX.findall(text)
    return [t for t in tokens if is_acronym(t)]

# -------------------------------------------------
# Helpers
# -------------------------------------------------
def build_breadcrumb(row: pd.Series) -> str:
    parts = []
    for col in LEVEL_COLUMNS:
        v = row.get(col)
        if pd.notna(v):
            s = str(v).strip()
            if s:
                parts.append(s)
    return " > ".join(parts)

# -------------------------------------------------
# Core analysis
# -------------------------------------------------
def analyze_file(df: pd.DataFrame):
    missing = [c for c in LEVEL_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing breadcrumb columns.\nExpected: Level 1 to Level 11\nMissing: {missing}\nFound: {list(df.columns)}"
        )

    rows_scanned = 0
    cells_scanned = 0

    out_rows = []
    acronym_counter = Counter()

    for idx, row in df.iterrows():
        rows_scanned += 1
        excel_row = idx + 2  # header in row 1
        breadcrumb = build_breadcrumb(row)

        for col in LEVEL_COLUMNS:
            v = row.get(col)
            if pd.isna(v):
                continue

            cell_value = str(v).strip()
            if not cell_value:
                continue

            cells_scanned += 1
            cell_address = f"{EXCEL_COLUMN_MAP[col]}{excel_row}"

            acronyms = extract_acronyms(cell_value)
            if not acronyms:
                continue

            for a in acronyms:
                acronym_counter[a] += 1
                out_rows.append({
                    "acronym": a,
                    "cell_address": cell_address,
                    "cell_value": cell_value,
                    "breadcrumb": breadcrumb
                })

    instances_df = pd.DataFrame(out_rows)
    if not instances_df.empty:
        instances_df = instances_df.drop_duplicates(subset=["acronym", "cell_address"])

    summary_df = pd.DataFrame(
        acronym_counter.most_common(),
        columns=["acronym", "count"]
    )

    metrics = {
        "rows_scanned": rows_scanned,
        "cells_scanned": cells_scanned,
        "unique_acronyms": int(summary_df.shape[0]),
        "unique_instances": int(instances_df.shape[0]),
    }

    return summary_df, instances_df, metrics

# -------------------------------------------------
# Sidebar controls
# -------------------------------------------------
with st.sidebar:
    st.subheader("Controls")
    uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    run = st.button("Run scan", type="primary", use_container_width=True)

    with st.expander("What counts as an acronym"):
        st.markdown(
            "- Two or more uppercase letters anywhere (e.g., **WiFi**, **USB**)\n"
            "- Uppercase + digit (e.g., **5G**, **3D**, **G4S**)\n"
            "- Subscript/superscript (e.g., **H₂O**, **m²**)\n"
            "- Unit-prefix CamelCase (e.g., **mmWave**, **µmSize**)\n"
            "- CamelCase terms (e.g., **iPhone**, **iPad**, **eSIM**)"
        )

# -------------------------------------------------
# Main flow
# -------------------------------------------------
if not uploaded_file:
    st.info("Step 1: Upload your Excel file from the sidebar.")
    st.stop()

try:
    df_raw = pd.read_excel(uploaded_file)
    df = normalize_columns(df_raw)
except Exception as e:
    st.error(f"Could not read the Excel file. Details: {e}")
    st.stop()

with st.expander("Preview uploaded data"):
    st.write(f"Columns detected: {len(df.columns)}")
    st.dataframe(df.head(15), use_container_width=True)

if not run:
    st.info("Click **Run scan** in the sidebar to analyze the file.")
    st.stop()

try:
    with st.spinner("Analyzing…"):
        summary_df, instances_df, metrics = analyze_file(df)
except Exception as e:
    st.error(str(e))
    st.stop()

# Metrics row
m1, m2, m3, m4 = st.columns(4)
m1.metric("Rows scanned", f"{metrics['rows_scanned']:,}")
m2.metric("Cells scanned", f"{metrics['cells_scanned']:,}")
m3.metric("Unique acronyms", f"{metrics['unique_acronyms']:,}")
m4.metric("Unique instances", f"{metrics['unique_instances']:,}")

st.divider()

tab1, tab2 = st.tabs(["📍 Instances", "⬇️ Download"])

with tab1:
    st.subheader("Acronyms with cell addresses")
    st.dataframe(instances_df, use_container_width=True, height=650)

with tab2:
    st.subheader("Download results")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        instances_df.to_excel(writer, index=False, sheet_name="Instances")
    output.seek(0)

    st.download_button(
        "Download results (Excel)",
        data=output,
        file_name="breadcrumb_acronyms_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
