import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Automation", layout="wide")

st.title("📊 Excel Automation Tool")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    try:
        # Read Excel file into pandas DataFrame
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        st.success("File uploaded and read successfully!")
        st.subheader("📄 Raw Data")
        st.dataframe(df, use_container_width=True)

        # -------------------------------------
        # 🧠 Basic Automation: Add totals column
        # -------------------------------------
        numeric_columns = df.select_dtypes(include="number").columns

        if not numeric_columns.empty:
            df["Total"] = df[numeric_columns].sum(axis=1)
            st.subheader("✅ Processed Data with Total")
            st.dataframe(df, use_container_width=True)
        else:
            st.warning("No numeric columns found to sum.")

        # -------------------------------------
        # 💾 Create downloadable Excel
        # -------------------------------------
        def to_excel(dataframe):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                dataframe.to_excel(writer, index=False, sheet_name='Processed')
            return output.getvalue()

        excel_bytes = to_excel(df)

        st.download_button(
            label="📥 Download Processed Excel",
            data=excel_bytes,
            file_name="processed_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
else:
    st.info("Please upload an Excel file to begin.")
