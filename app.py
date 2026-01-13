import re
from collections import Counter
from io import BytesIO

import pandas as pd
import streamlit as st

# -------------------------------------------------
# App configuration
# -------------------------------------------------
st.set_page_config(page_title="Breadcrumb Acronym Finder", layout="wide")

st.title("Breadcrumb Acronym Finder")
st.caption(
    "Uploads an Excel breadcrumb file, scans all breadcrumb columns "
    "B–L (Level 1–11), and returns acronyms with exact Excel cell addresses."
)

# -------------------------------------------------
# Acronym rules (FINAL, EXTENDED)
# -------------------------------------------------

# Unicode subscript and superscript digits
SUBSCRIPT_DIGITS = set("₀₁₂₃₄₅₆₇₈₉")
SUPERSCRIPT_DIGITS = set("⁰¹²³⁴⁵⁶⁷⁸⁹²³¹")

# Common unit / metric prefixes (for mmWave, µmSize, etc.)
UNIT_PREFIXES = {
    "mm", "cm", "km", "nm", "pm",
    "um", "µm", "μm"
}

# Candidate token pattern:
# letters, digits, sub/superscripts with internal separators
CANDIDATE_TOKEN_REGEX = re.compile(
    r"[A-Za-z0-9₀₁₂₃₄₅₆₇₈₉⁰¹²³⁴⁵⁶⁷⁸⁹²³¹]+(?:[-_/.][A-Za-z0-9₀₁₂₃₄₅₆₇₈₉⁰¹²³⁴⁵⁶⁷⁸⁹²³¹]+)*"
)


def is_camel_case(token: str) -> bool:
    """
    Detect CamelCase or mixed-case words like:
    iPhone, eSIM, PowerBankPro
    Excludes normal capitalized words like Construction.
    """
    if token[0].isupper() and token[1:].islower():
        return False  # Normal capitalized word

    has_lower = any(ch.islower() for ch in token)
    has_upper = any(ch.isupper() for ch in token)

    return has_lower and has_upper


def is_acronym(token: str) -> bool:
    """
    Acronym / special-term rules:
    """
    if not token or len(token) < 2:
        return False

    uppercase_count = sum(1 for ch in token if "A" <= ch <= "Z")
    has_upper = uppercase_count > 0
    has_digit = any(ch.isdigit() for ch in token)
    has_sub_or_sup = any(
        (ch in SUBSCRIPT_DIGITS) or (ch in SUPERSCRIPT_DIGITS) for ch in token
    )
    has_letter = any(ch.isalpha() for ch in token)

    # Rule 1: two or more uppercase letters anywhere
    if uppercase_count >= 2:
        return True

    # Rule 2: uppercase letter + digit
    if has_upper and has_digit:
        return True

    # Rule 3: subscript/superscript + letter
    if has_sub_or_sup and has_letter:
        return True

    # Rule 4: unit-prefix + CamelCase (mmWave, µmSize)
    lowered = token.lower()
    for p in UNIT_PREFIXES:
        if lowered.startswith(p.lower()):
            rest = token[len(p):]
            if any("A" <= ch <= "Z" for ch in rest):
                return True

    # Rule 5: general CamelCase words (iPhone, iPad, eSIM)
    if is_camel_case(token):
        return True

    return False


def extract_acronyms(text: str):
    if not isinstance(text, str) or not text.strip():
        return []
    tokens = CANDIDATE_TOKEN_REGEX.findall(text)
    return [t for t in tokens if is_acronym(t)]


# -------------------------------------------------
# Expected breadcrumb columns (Excel B–L)
# -------------------------------------------------
LEVEL_COLUMNS = [
    "Level 1", "Level 2", "Level 3", "Level 4", "Level 5",
    "Level 6", "Level 7", "Level  8", "Level 9", "Level 10", "Level 11"
]

EXCEL_COLUMN_MAP = {
    "Level 1": "B",
    "Level 2": "C",
    "Level 3": "D",
    "Level 4": "E",
    "Level 5": "F",
    "Level 6": "G",
    "Level 7": "H",
    "Level  8": "I",
    "Level 9": "J",
    "Level 10": "K",
    "Level 11": "L",
}


# -------------------------------------------------
# Helper: build breadcrumb
# -------------------------------------------------
def build_breadcrumb(row):
    parts = []
    for col in LEVEL_COLUMNS:
        val = row.get(col)
        if pd.notna(val):
            val = str(val).strip()
            if val:
                parts.append(val)
    return " > ".join(parts)


# -------------------------------------------------
# Core analysis
# -------------------------------------------------
def analyze_all_rows(df: pd.DataFrame):
    missing = [c for c in LEVEL_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "Missing breadcrumb columns.\n"
            "Expected: Level 1 to Level 11 (Excel B–L)\n"
            f"Missing: {missing}\n"
            f"Found columns: {list(df.columns)}"
        )

    acronym_counter = Counter()
    rows_out = []

    for idx, row in df.iterrows():
        excel_row_number = idx + 2
        breadcrumb = build_breadcrumb(row)

        for col in LEVEL_COLUMNS:
            cell_value = row.get(col)
            if pd.isna(cell_value):
                continue

            cell_value = str(cell_value).strip()
            if not cell_value:
                continue

            matches = extract_acronyms(cell_value)
            if not matches:
                continue

            cell_address = f"{EXCEL_COLUMN_MAP[col]}{excel_row_number}"

            for acronym in matches:
                acronym_counter[acronym] += 1
                rows_out.append({
                    "acronym": acronym,
                    "cell_address": cell_address,
                    "cell_value": cell_value,
                    "breadcrumb": breadcrumb
                })

    instances_df = pd.DataFrame(rows_out)
    if not instances_df.empty:
        instances_df = instances_df.drop_duplicates(
            subset=["acronym", "cell_address"]
        )

    summary_df = pd.DataFrame(
        acronym_counter.most_common(),
        columns=["acronym", "count"]
    )

    return summary_df, instances_df


# -------------------------------------------------
# UI
# -------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload your breadcrumb Excel file (.xlsx)",
    type=["xlsx"]
)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        st.subheader("Analyze entire file")
        st.write(
            "This will scan **every row** and **all breadcrumb columns "
            "B–L (Level 1–11)**."
        )

        if st.button("Analyze file"):
            summary_df, instances_df = analyze_all_rows(df)

            st.success(
                f"Done. Unique acronyms: {len(summary_df)} | "
                f"Unique instances: {len(instances_df)}"
            )

            st.subheader("Summary – Acronym counts")
            st.dataframe(summary_df, use_container_width=True, height=350)

            st.subheader("Instances – Exact cell addresses")
            st.dataframe(instances_df, use_container_width=True, height=450)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                summary_df.to_excel(writer, index=False, sheet_name="Summary")
                instances_df.to_excel(writer, index=False, sheet_name="Instances")
            output.seek(0)

            st.download_button(
                label="Download results (Excel)",
                data=output,
                file_name="breadcrumb_acronyms_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(str(e))
else:
    st.info("Upload an Excel (.xlsx) breadcrumb file to begin.")
