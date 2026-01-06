import pandas as pd
import streamlit as st
from io import BytesIO

# -------------------------------
# Utilities
# -------------------------------

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )
    return df


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean data WITHOUT corrupting types.
    Missing values remain NaN internally.
    """
    df = normalize_columns(df)

    # Strip strings and normalize empty strings → NaN
    for col in df.select_dtypes(include="object"):
        df[col] = (
            df[col]
            .str.strip()
            .replace("", pd.NA)
        )

    # Drop full-row duplicates
    df = df.drop_duplicates()

    return df


def validate_data(df: pd.DataFrame, required_columns: list[str]) -> list[str]:
    issues = []

    # Duplicate rows
    dup_count = df.duplicated().sum()
    if dup_count > 0:
        issues.append(f"{dup_count} duplicate rows found")

    # Required columns
    for col in required_columns:
        if col not in df.columns:
            issues.append(f"Missing required column: {col}")
        else:
            missing_count = df[col].isna().sum()
            if missing_count > 0:
                issues.append(f"Column '{col}' has {missing_count} missing values")

    return issues


def display_safe(df: pd.DataFrame) -> pd.DataFrame:
    """
    UI-only representation.
    Converts NaN → 'MISSING' and casts to string
    to avoid Arrow / Streamlit failures.
    """
    return df.fillna("MISSING").astype(str)


# -------------------------------
# Streamlit UI
# -------------------------------

st.set_page_config(page_title="Excel Cleaner", layout="wide")
st.title("Excel Data Cleaning & Validation Tool")

uploaded_file = st.file_uploader("Upload Excel file (single sheet, headers required)", type=["xlsx"])

if "proceed" not in st.session_state:
    st.session_state.proceed = False

if uploaded_file:
    raw_df = pd.read_excel(uploaded_file)

    st.subheader("Raw Data Preview")
    st.dataframe(display_safe(raw_df.head()))

    if st.button("Proceed with Cleaning & Validation"):
        st.session_state.proceed = True

if st.session_state.proceed:
    # Normalize once, everywhere
    raw_df = normalize_columns(raw_df)

    st.subheader("Select Required Columns")
    required_columns = st.multiselect(
        "Required columns",
        options=list(raw_df.columns),
        default=[c for c in ["date", "amount"] if c in raw_df.columns]
    )

    # -------------------------------
    # Validation (RAW but normalized)
    # -------------------------------
    issues = validate_data(raw_df, required_columns)

    st.subheader("Validation Results")
    if issues:
        for issue in issues:
            st.warning(issue)
    else:
        st.success("No validation issues found")

    # -------------------------------
    # Cleaning
    # -------------------------------
    cleaned_df = clean_data(raw_df)

    st.subheader("Cleaned Data Preview")
    st.dataframe(display_safe(cleaned_df.head()))

    # -------------------------------
    # Excel Output
    # -------------------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cleaned_df.to_excel(writer, index=False, sheet_name="Cleaned_Data")

        report = (
            pd.DataFrame({"Validation_Issues": issues})
            if issues
            else pd.DataFrame({"Validation_Issues": ["No issues found"]})
        )
        report.to_excel(writer, index=False, sheet_name="Validation_Report")

    st.download_button(
        "Download Cleaned Excel with Validation Report",
        data=output.getvalue(),
        file_name="cleaned_data_with_validation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
