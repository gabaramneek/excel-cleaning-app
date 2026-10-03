import streamlit as st
import pandas as pd
import re
from io import BytesIO


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Excel Data Cleaner",
    page_icon="🧹",
    layout="wide"
)


# ============================================================
# TITLE
# ============================================================

st.title("🧹 Excel Data Cleaning Automation")

st.write(
    "Upload an Excel file and automatically clean, "
    "validate and analyze your data."
)


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "Upload your Excel file",
    type=["xlsx", "xls"]
)


# ============================================================
# CLEANING FUNCTION
# ============================================================

def clean_data(df):

    quality_report = []

    def add_issue(row_number, column, issue, value):
        quality_report.append({
            "Row": row_number,
            "Column": column,
            "Issue": issue,
            "Value": value
        })

    original_rows = len(df)

    # --------------------------------------------------------
    # Clean column names
    # --------------------------------------------------------

    df.columns = (
        df.columns
        .str.strip()
        .str.replace(" ", "_")
    )

    # --------------------------------------------------------
    # Text columns
    # --------------------------------------------------------

    text_columns = [
        "Name",
        "Email",
        "City",
        "State",
        "Product",
        "Category",
        "Payment_Method"
    ]

    for col in text_columns:

        if col in df.columns:

            df[col] = (
                df[col]
                .astype("string")
                .str.strip()
            )

    # --------------------------------------------------------
    # Names
    # --------------------------------------------------------

    if "Name" in df.columns:

        df["Name"] = (
            df["Name"]
            .str.replace(r"\s+", " ", regex=True)
            .str.title()
        )

    # --------------------------------------------------------
    # City
    # --------------------------------------------------------

    if "City" in df.columns:

        df["City"] = (
            df["City"]
            .str.replace(r"\s+", " ", regex=True)
            .str.title()
        )

    # --------------------------------------------------------
    # State
    # --------------------------------------------------------

    if "State" in df.columns:

        df["State"] = (
            df["State"]
            .str.replace(r"\s+", " ", regex=True)
            .str.title()
        )

    # --------------------------------------------------------
    # Product
    # --------------------------------------------------------

    if "Product" in df.columns:

        df["Product"] = (
            df["Product"]
            .str.replace(r"\s+", " ", regex=True)
            .str.title()
        )

    # --------------------------------------------------------
    # Category
    # --------------------------------------------------------

    if "Category" in df.columns:

        df["Category"] = (
            df["Category"]
            .str.replace(r"\s+", " ", regex=True)
            .str.title()
        )

    # --------------------------------------------------------
    # Payment Method
    # --------------------------------------------------------

    if "Payment_Method" in df.columns:

        df["Payment_Method"] = (
            df["Payment_Method"]
            .str.replace(r"\s+", " ", regex=True)
            .str.title()
        )

    # --------------------------------------------------------
    # Email
    # --------------------------------------------------------

    if "Email" in df.columns:

        df["Email"] = df["Email"].str.lower()

        email_pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"

        def validate_email(email):

            if pd.isna(email) or email == "":
                return "Missing"

            if re.match(email_pattern, email):
                return "Valid"

            return "Invalid"

        df["Email_Status"] = df["Email"].apply(
            validate_email
        )

        for index, row in df.iterrows():

            if row["Email_Status"] == "Invalid":

                add_issue(
                    index + 2,
                    "Email",
                    "Invalid email address",
                    row["Email"]
                )

            elif row["Email_Status"] == "Missing":

                add_issue(
                    index + 2,
                    "Email",
                    "Missing email address",
                    ""
                )

    # --------------------------------------------------------
    # Phone
    # --------------------------------------------------------

    if "Phone" in df.columns:

        def clean_phone(phone):

            if pd.isna(phone):
                return pd.NA

            digits = re.sub(
                r"\D",
                "",
                str(phone)
            )

            if (
                digits.startswith("91")
                and len(digits) == 12
            ):
                digits = digits[2:]

            if len(digits) == 10:
                return "+91" + digits

            return pd.NA

        df["Phone"] = df["Phone"].apply(
            clean_phone
        )

    # --------------------------------------------------------
    # Dates
    # --------------------------------------------------------

    if "Order_Date" in df.columns:

        df["Order_Date"] = pd.to_datetime(
            df["Order_Date"],
            errors="coerce",
            dayfirst=True
        )

    # --------------------------------------------------------
    # Numeric columns
    # --------------------------------------------------------

    if "Quantity" in df.columns:

        df["Quantity"] = pd.to_numeric(
            df["Quantity"],
            errors="coerce"
        )

    if "Unit_Price" in df.columns:

        df["Unit_Price"] = pd.to_numeric(
            df["Unit_Price"],
            errors="coerce"
        )

    # --------------------------------------------------------
    # Negative values
    # --------------------------------------------------------

    if "Quantity" in df.columns:

        for index, row in df.iterrows():

            if (
                pd.notna(row["Quantity"])
                and row["Quantity"] < 0
            ):

                add_issue(
                    index + 2,
                    "Quantity",
                    "Negative quantity",
                    row["Quantity"]
                )

    if "Unit_Price" in df.columns:

        for index, row in df.iterrows():

            if (
                pd.notna(row["Unit_Price"])
                and row["Unit_Price"] < 0
            ):

                add_issue(
                    index + 2,
                    "Unit_Price",
                    "Negative price",
                    row["Unit_Price"]
                )

    # --------------------------------------------------------
    # Duplicate records
    # --------------------------------------------------------

    duplicates_removed = 0

    if "Customer_ID" in df.columns:

        duplicates_removed = df.duplicated(
            subset=["Customer_ID"],
            keep="first"
        ).sum()

        df = df.drop_duplicates(
            subset=["Customer_ID"],
            keep="first"
        )

    # --------------------------------------------------------
    # Issues dataframe
    # --------------------------------------------------------

    issues_df = pd.DataFrame(
        quality_report
    )

    return (
        df,
        issues_df,
        original_rows,
        duplicates_removed
    )


# ============================================================
# PROCESS FILE
# ============================================================

if uploaded_file is not None:

    df = pd.read_excel(
        uploaded_file
    )

    st.success("Excel file uploaded successfully!")

    # --------------------------------------------------------
    # ORIGINAL DATA
    # --------------------------------------------------------

    st.subheader("Original Data")

    st.dataframe(
        df.head(10),
        use_container_width=True
    )

    # --------------------------------------------------------
    # CLEAN BUTTON
    # --------------------------------------------------------

    if st.button(
        "🧹 Clean Data",
        type="primary"
    ):

        (
            cleaned_df,
            issues_df,
            original_rows,
            duplicates_removed
        ) = clean_data(df)

        # Store results
        st.session_state["cleaned_df"] = cleaned_df
        st.session_state["issues_df"] = issues_df
        st.session_state["original_rows"] = original_rows
        st.session_state["duplicates_removed"] = duplicates_removed


# ============================================================
# DISPLAY RESULTS
# ============================================================

if "cleaned_df" in st.session_state:

    cleaned_df = st.session_state["cleaned_df"]
    issues_df = st.session_state["issues_df"]

    original_rows = st.session_state[
        "original_rows"
    ]

    duplicates_removed = st.session_state[
        "duplicates_removed"
    ]

    st.divider()

    st.subheader("📊 Data Quality Summary")

    col1, col2, col3, col4 = st.columns(4)

    with col1:

        st.metric(
            "Original Rows",
            original_rows
        )

    with col2:

        st.metric(
            "Cleaned Rows",
            len(cleaned_df)
        )

    with col3:

        st.metric(
            "Duplicates Removed",
            duplicates_removed
        )

    with col4:

        st.metric(
            "Issues Detected",
            len(issues_df)
        )

    # --------------------------------------------------------
    # CLEANED DATA
    # --------------------------------------------------------

    st.subheader("✅ Cleaned Data")

    st.dataframe(
        cleaned_df,
        use_container_width=True
    )

    # --------------------------------------------------------
    # ISSUES
    # --------------------------------------------------------

    st.subheader("⚠️ Data Quality Issues")

    if len(issues_df) > 0:

        st.dataframe(
            issues_df,
            use_container_width=True
        )

    else:

        st.success(
            "No data quality issues detected!"
        )

    # --------------------------------------------------------
    # DOWNLOAD
    # --------------------------------------------------------

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        cleaned_df.to_excel(
            writer,
            sheet_name="Cleaned_Data",
            index=False
        )

        issues_df.to_excel(
            writer,
            sheet_name="Data_Quality_Issues",
            index=False
        )

    output.seek(0)

    st.download_button(
        label="⬇️ Download Cleaned Excel",
        data=output,
        file_name="cleaned_customer_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )