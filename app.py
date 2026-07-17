import streamlit as st
from io import BytesIO

from converters.expense import convert as expense_convert
from converters.benefits import convert as benefits_convert
from converters.sales import convert as sales_convert

st.set_page_config(
    page_title="Budget Database Toolkit",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Budget Database Toolkit")

converter = st.selectbox(
    "Choose a converter",
    [
        "Expense Template",
        "Benefits Template",
        "Sales Template"

    ]
)

uploaded_file = st.file_uploader(
    "Upload Excel Workbook",
    type=["xlsx"]
)

if uploaded_file:

    if st.button("Convert"):

        with st.spinner("Processing workbook..."):

            if converter == "Expense Template":
                df = expense_convert(uploaded_file)

            elif converter == "Benefits Template":
                df = benefits_convert(uploaded_file)

            elif converter == "Sales Template":
                df = sales_convert(uploaded_file)

        st.success(f"{len(df):,} rows created.")

        st.dataframe(df.head(100), use_container_width=True)

        output = BytesIO()

        df.to_excel(output, index=False)

        output.seek(0)

        st.download_button(
            "Download Database",
            output,
            "Database.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )