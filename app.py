import streamlit as st
import pandas as pd

# ============================================
# PAGE CONFIG
# ============================================

st.set_page_config(
    page_title="Insurance Premium Calculator",
    page_icon="💰",
    layout="wide"
)

# ============================================
# TITLE
# ============================================

st.title("💰 Group Term Life Insurance Premium Calculator")
st.markdown("Quick calculator or upload an Excel file for premium calculation.")

# ============================================
# SETTINGS
# ============================================

RATE_PER_LAKH = 435          # Inclusive of GST
GST_RATE = 0.18

# ============================================
# QUICK PREMIUM CALCULATOR
# ============================================

st.subheader("⚡ Quick Premium Calculator")

quick_sum_assured = st.number_input(
    "Enter Sum Assured (₹)",
    min_value=0.0,
    value=1000000.0,
    step=100000.0,
    key="quick_sa"
)

if st.button("Calculate Premium"):

    total_premium = (quick_sum_assured / 100000) * RATE_PER_LAKH
    premium_excl = total_premium / (1 + GST_RATE)
    gst_amount = total_premium - premium_excl

    c1, c2, c3 = st.columns(3)

    c1.metric("Premium Excl GST", f"₹ {premium_excl:,.2f}")
    c2.metric("GST (18%)", f"₹ {gst_amount:,.2f}")
    c3.metric("Total Premium (Incl GST)", f"₹ {total_premium:,.2f}")

st.divider()

# ============================================
# FILE UPLOAD
# ============================================

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

# ============================================
# PROCESS FILE
# ============================================

if uploaded_file is not None:

    try:

        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()

        st.subheader("Uploaded Data")
        st.dataframe(df.head())

        required_columns = [
            'Loan Account No.',
            'Name of Primary Loan borrower',
            'Mobile No',
            'Sum Assured'
        ]

        for col in required_columns:
            if col not in df.columns:
                df[col] = ""

        df['Sum Assured'] = pd.to_numeric(
            df['Sum Assured'],
            errors='coerce'
        ).fillna(0)

        if 'MAIN MEMBER AGE' not in df.columns:
            df['MAIN MEMBER AGE'] = 0

        if 'Loan Outstanding Amount' not in df.columns:
            df['Loan Outstanding Amount'] = 0

        # Total Premium (Inclusive of GST)
        df['Premium + GST'] = (
            (df['Sum Assured'] / 100000)
            * RATE_PER_LAKH
        )

        # Premium Excluding GST
        df['Premium Excl GST'] = (
            df['Premium + GST'] / (1 + GST_RATE)
        )

        # GST Amount
        df['GST Amount'] = (
            df['Premium + GST'] - df['Premium Excl GST']
        )

        output_columns = [
            'Loan Account No.',
            'Name of Primary Loan borrower',
            'Mobile No',
            'MAIN MEMBER AGE',
            'Sum Assured',
            'Premium Excl GST',
            'GST Amount',
            'Premium + GST'
        ]

        final_df = df[output_columns]

        st.subheader("Portfolio Summary")

        col1, col2, col3 = st.columns(3)

        col1.metric("Total Members", len(final_df))
        col2.metric(
            "Total Sum Assured",
            f"₹ {final_df['Sum Assured'].sum():,.0f}"
        )
        col3.metric(
            "Total Premium",
            f"₹ {final_df['Premium + GST'].sum():,.2f}"
        )

        st.subheader("Premium Calculation Output")
        st.dataframe(final_df, use_container_width=True)

        output_file = "Premium_Output.xlsx"
        final_df.to_excel(output_file, index=False)

        with open(output_file, "rb") as file:
            st.download_button(
                label="⬇ Download Output Excel",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
