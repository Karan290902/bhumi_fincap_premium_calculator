import io
import re
import pandas as pd
import streamlit as st

# --- CONSTANTS ---
RATE_PER_LAKH = 435.0
GST_RATE = 0.18

# Define the exact required output columns in order
OUTPUT_COLUMNS = [
    "Sr. No.",
    "Customer ID",
    "Name of the Person covered",
    "Gender (M/F)",
    "DOB",
    "Member Date of Joining in the Organization",
    "Effective date of Coverage",
    "Opted Sum Assured",
    "Designation",
    "Occupation",
    "Annual Salary",
    "Marital Status",
    "Nominee Name",
    "Nominee DOB",
    "Nominee Gender",
    "Relationship of the Nominee with Insurance covered Person",
    "Member Contact Number Mandatory",
    "Member Email ID Mandatory",
    "Zone",
    "Branch Name",
    "Branch Code",
    "Type of Age Proof",
    "Address",
    "City",
    "State",
    "Pin code",
    "Premium",
    "GST",
    "Total Premium",
]


# --- HELPER FUNCTIONS ---
def normalize_text(text: str) -> str:
    """Normalize headers for flexible matching by removing line breaks,

    extra spaces, and converting to lowercase.
    """
    if not isinstance(text, str):
        return ""
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()


def calculate_premium_metrics(opted_sum_assured: float):
    """Calculate Total Premium, Premium (excl. GST), and GST components."""
    total_premium = (opted_sum_assured / 100000.0) * RATE_PER_LAKH
    premium = total_premium / (1.0 + GST_RATE)
    gst = total_premium - premium
    return premium, gst, total_premium


# --- APP SETUP ---
st.set_page_config(
    page_title="Group Term Life Insurance Premium Calculator", layout="wide"
)
st.title("Group Term Life (GTL) Insurance Premium Calculator")

# Sidebar navigation between modules
module = st.sidebar.radio("Select Module", ["Quick Calculator", "Bulk Calculator"])

# --- MODULE 1: QUICK PREMIUM CALCULATOR ---
if module == "Quick Calculator":
    st.header("Quick Premium Calculator")
    st.write("Calculate premium for an individual instantly.")

    opted_sa = st.number_input(
        "Opted Sum Assured (₹)", min_value=0.0, step=50000.0, value=1000000.0
    )

    if opted_sa > 0:
        prem, gst_val, total_prem = calculate_premium_metrics(opted_sa)

        col1, col2, col3 = st.columns(3)
        col1.metric("Premium (Excl. GST)", f"₹{prem:,.2f}")
        col2.metric("GST (18%)", f"₹{gst_val:,.2f}")
        col3.metric("Total Premium (Incl. GST)", f"₹{total_prem:,.2f}")
    else:
        st.warning("Please enter a Sum Assured greater than 0.")

# --- MODULE 2: BULK PREMIUM CALCULATOR ---
else:
    st.header("Bulk Premium Calculator")
    st.write("Upload an Excel file to calculate premiums for multiple members.")

    uploaded_file = st.file_uploader(
        "Upload Excel File (.xlsx)", type=["xlsx"]
    )

    if uploaded_file is not None:
        try:
            # Read input Excel file
            df_input = pd.read_excel(uploaded_file, engine="openpyxl")

            # Store mapping of normalized column name to original column name
            norm_to_orig = {
                normalize_text(col): col for col in df_input.columns
            }

            # Mandatory target columns and their normalized targets
            target_name_norm = normalize_text("Name of the Person covered")
            target_sa_norm = normalize_text("Opted Sum Assured")

            # Check presence of mandatory columns
            missing_cols = []
            if target_name_norm not in norm_to_orig:
                missing_cols.append("• Name of the Person covered")
            if target_sa_norm not in norm_to_orig:
                missing_cols.append("• Opted Sum Assured")

            if missing_cols:
                st.error("Missing mandatory columns:\n" + "\n".join(missing_cols))
                st.stop()

            # Identify the actual original column names in the uploaded sheet
            orig_name_col = norm_to_orig[target_name_norm]
            orig_sa_col = norm_to_orig[target_sa_norm]

            # Build the template output DataFrame
            df_output = pd.DataFrame(columns=OUTPUT_COLUMNS)

            # Map existing columns from the uploaded file to the strict output format
            for col in OUTPUT_COLUMNS:
                norm_output_col = normalize_text(col)
                if norm_output_col in norm_to_orig:
                    # Direct case/space-insensitive match exists
                    df_output[col] = df_input[norm_to_orig[norm_output_col]]
                elif norm_output_col == target_name_norm:
                    df_output[col] = df_input[orig_name_col]
                elif norm_output_col == target_sa_norm:
                    df_output[col] = df_input[orig_sa_col]
                else:
                    # Does not exist, leave blank
                    df_output[col] = None

            # Enforce data type for numeric calculation
            df_output["Opted Sum Assured"] = (
                pd.to_numeric(df_output["Opted Sum Assured"], errors="coerce")
                .fillna(0)
                .astype(float)
            )

            # Calculate Premiums row by row
            premiums = []
            gsts = []
            total_premiums = []

            for sa in df_output["Opted Sum Assured"]:
                p, g, tp = calculate_premium_metrics(sa)
                premiums.append(p)
                gsts.append(g)
                total_premiums.append(tp)

            df_output["Premium"] = premiums
            df_output["GST"] = gsts
            df_output["Total Premium"] = total_premiums

            # Populate Serial Number if it's blank or missing
            df_output["Sr. No."] = range(1, len(df_output) + 1)

            # Re-ensure exact column alignment and ordering
            df_output = df_output[OUTPUT_COLUMNS]

            # --- PORTFOLIO SUMMARY ---
            st.subheader("Portfolio Summary")
            sum_col1, sum_col2, sum_col3 = st.columns(3)

            total_members = len(df_output)
            total_sa = df_output["Opted Sum Assured"].sum()
            total_prem_sum = df_output["Total Premium"].sum()

            sum_col1.metric("Total Members", f"{total_members:,}")
            sum_col2.metric("Total Sum Assured", f"₹{total_sa:,.2f}")
            sum_col3.metric("Total Premium (Incl. GST)", f"₹{total_prem_sum:,.2f}")

            # --- OUTPUT PREVIEW ---
            st.subheader("Output Preview")
            st.dataframe(df_output)

            # --- EXPORT TO EXCEL ---
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(
                output_buffer, engine="openpyxl"
            ) as writer:
                df_output.to_excel(writer, index=False, sheet_name="GTL Premium")
            excel_data = output_buffer.getvalue()

            st.download_button(
                label="Download Output Excel",
                data=excel_data,
                file_name="Bhumi_GTL_Premium_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")