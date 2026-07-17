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

st.write(
    "Convert budgeting templates into Power BI-ready databases."
)

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
                filename = "Expense_Database.xlsx"

            elif converter == "Benefits Template":
                df = benefits_convert(uploaded_file)
                filename = "Benefits_Database.xlsx"

            elif converter == "Sales Template":
                df = sales_convert(uploaded_file)
                filename = "Sales_Database.xlsx"

        st.success(f"Finished! {len(df):,} records created.")

        st.dataframe(
            df.head(100),
            use_container_width=True
        )

        output = BytesIO()

        df.to_excel(
            output,
            index=False,
            engine="openpyxl"
        )

        output.seek(0)

        st.download_button(
            label="📥 Download Database",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
