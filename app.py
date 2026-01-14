import re
from collections import Counter
from io import BytesIO

import pandas as pd
import streamlit as st

# -------------------------------------------------
# App configuration
# -------------------------------------------------
st.set_page_config(page_title="Breadcrumb Acronym Integrity (Meta-00)", layout="wide")

st.title("Breadcrumb Acronym Integrity Checker (Meta-00)")
st.caption(
    "Uploads an Excel breadcrumb file, scans breadcrumb columns (Level 1–11), "
    "detects acronyms and corrupted acronyms (e.g., SMS → Sms), and returns "
    "exact Excel cell addresses with fix suggestions."
)

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
# Normalize column headers (CRITICAL FIX)
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
# Acronym detection rules (Meta-00)
# -------------------------------------------------
SUBSCRIPT_DIGITS = set("₀₁₂₃₄₅₆₇₈₉")
SUPERSCRIPT_DIGITS = set("⁰¹²³⁴⁵⁶⁷⁸⁹¹²³")
UNIT_PREFIXES = {"mm", "cm", "km", "nm", "pm", "um", "µm", "μm"}

KNOWN_ACRONYMS = {
    "SMS", "MMS", "API", "SDK", "CRM", "ERP", "POS", "SKU",
    "B2B", "B2C", "C2C", "D2C",
    "AI", "ML", "NLP", "LLM",
    "SEO", "SEM", "PPC", "KPI", "OKR", "ROI",
    "UX", "UI", "QA", "SLA", "ETA",
    "OTP", "PIN", "VPN", "LAN", "WAN", "WIFI", "WIFi", "WiFi",
    "GPS", "RFID", "NFC", "SIM", "ESIM",
    "PDF", "CSV", "XML", "JSON",
    "HR", "IT", "BI",
}

COMMON_TITLECASE_WORDS = {
    "And", "Or", "The", "A", "An", "Of", "To", "In", "On", "For", "With",
    "Home", "Service", "Services", "Product", "Products", "Store", "Stores",
    "Repair", "Cleaning", "Marketing", "Design", "Support", "Management",
}

CANDIDATE_TOKEN_REGEX = re.compile(
    r"[A-Za-z0-9₀₁₂₃₄₅₆₇₈₉⁰¹²³⁴⁵⁶⁷⁸⁹¹²³]+(?:[-_/.][A-Za-z0-9₀₁₂₃₄₅₆₇₈₉⁰¹²³⁴⁵⁶⁷⁸⁹¹²³]+)*"
)

PASCAL_CASE_SHORT = re.compile(r"^[A-Z][a-z]{1,5}$")  # Sms, Api, Crm

# -------------------------------------------------
# Helper functions
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


def is_camel_case(token: str) -> bool:
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


def is_strong_acronym(token: str) -> bool:
    if not token or len(token) < 2:
        return False

    upper_count = sum(c.isupper() for c in token)
    has_digit = any(c.isdigit() for c in token)
    has_sub_sup = any(c in SUBSCRIPT_DIGITS or c in SUPERSCRIPT_DIGITS for c in token)

    if token.upper() in KNOWN_ACRONYMS:
        return True
    if upper_count >= 2:
        return True
    if any(c.isupper() for c in token) and has_digit:
        return True
    if has_sub_sup:
        return True
    if looks_like_unit_prefix_camel(token):
        return True
    if is_camel_case(token):
        return True

    return False


def is_corrupted_acronym(token: str, token_upper_seen: Counter) -> bool:
    if token in COMMON_TITLECASE_WORDS:
        return False
    if not PASCAL_CASE_SHORT.match(token):
        return False
    if token.upper() in KNOWN_ACRONYMS:
        return True
    if token_upper_seen[token.upper()] > 0 and token != token.upper():
        return True
    return False


def extract_tokens(text: str):
    if not isinstance(text, str):
        return []
    return CANDIDATE_TOKEN_REGEX.findall(text)

# -------------------------------------------------
# Core analysis
# -------------------------------------------------
def analyze_file(df: pd.DataFrame):
    missing = [c for c in LEVEL_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing breadcrumb columns.\n"
            f"Expected: Level 1 to Level 11\n"
            f"Missing: {missing}\n"
            f"Found: {list(df.columns)}"
        )

    token_upper_seen = Counter()

    for _, row in df.iterrows():
        for col in LEVEL_COLUMNS:
            v = row.get(col)
            if pd.notna(v):
                for t in extract_tokens(str(v)):
                    token_upper_seen[t.upper()] += 1

    detected, corrupted = [], []

    for idx, row in df.iterrows():
        row_num = idx + 2
        breadcrumb = build_breadcrumb(row)

        for col in LEVEL_COLUMNS:
            v = row.get(col)
            if pd.isna(v):
                continue

            cell_value = str(v).strip()
            if not cell_value:
                continue

            cell_address = f"{EXCEL_COLUMN_MAP[col]}{row_num}"

            for token in extract_tokens(cell_value):
                if is_strong_acronym(token):
                    detected.append({
                        "term": token,
                        "cell_address": cell_address,
                        "cell_value": cell_value,
                        "breadcrumb": breadcrumb
                    })

                if is_corrupted_acronym(token, token_upper_seen):
                    corrupted.append({
                        "corrupted_term": token,
                        "suggested_fix": token.upper(),
                        "cell_address": cell_address,
                        "cell_value": cell_value,
                        "breadcrumb": breadcrumb
                    })

    detected_df = pd.DataFrame(detected).drop_duplicates(["term", "cell_address"])
    corrupted_df = pd.DataFrame(corrupted).drop_duplicates(["corrupted_term", "cell_address"])

    detected_summary = pd.DataFrame(
        Counter(detected_df["term"]).most_common(),
        columns=["term", "count"]
    ) if not detected_df.empty else pd.DataFrame(columns=["term", "count"])

    corrupted_summary = pd.DataFrame(
        Counter(corrupted_df["corrupted_term"]).most_common(),
        columns=["corrupted_term", "count"]
    ) if not corrupted_df.empty else pd.DataFrame(columns=["corrupted_term", "count"])

    return detected_summary, detected_df, corrupted_summary, corrupted_df

# -------------------------------------------------
# UI
# -------------------------------------------------
uploaded_file = st.file_uploader("Upload your breadcrumb Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df = normalize_columns(df)

        if st.button("Analyze file"):
            ds, ddf, cs, cdf = analyze_file(df)

            st.success(f"Done. Corruptions found: {len(cdf)} | Terms detected: {len(ddf)}")

            st.subheader("Corrupted acronyms (must fix)")
            st.dataframe(cs, use_container_width=True, height=250)
            st.dataframe(cdf, use_container_width=True, height=450)

            st.subheader("Detected acronyms / special terms")
            st.dataframe(ds, use_container_width=True, height=250)
            st.dataframe(ddf, use_container_width=True, height=450)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                cs.to_excel(writer, index=False, sheet_name="Corruption_Summary")
                cdf.to_excel(writer, index=False, sheet_name="Corruptions")
                ds.to_excel(writer, index=False, sheet_name="Detected_Summary")
                ddf.to_excel(writer, index=False, sheet_name="Detected_Terms")
            output.seek(0)

            st.download_button(
                "Download results (Excel)",
                output,
                "breadcrumb_acronym_integrity_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(str(e))
else:
    st.info("Upload an Excel (.xlsx) breadcrumb file to begin.")
