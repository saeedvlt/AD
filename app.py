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
st.caption("Convert budget templates into clean databases.")

st.info(
    """
    1. Select the template type.
    2. Upload the Excel workbook.
    3. Click Download.
    """
)


CONVERTERS = {
    "Expense Template": (
        expense_convert,
        "Expense Database.xlsx",
    ),
    "Benefits Template": (
        benefits_convert,
        "Benefits Database.xlsx",
    ),
    "Sales Template": (
        sales_convert,
        "Sales Database.xlsx",
    ),
}

converter = st.selectbox(
    "Choose a converter",
    list(CONVERTERS.keys())
)

upload_labels = {
    "Expense Template": "Upload Expense Template Workbook",
    "Benefits Template": "Upload Benefits Template Workbook",
    "Sales Template": "Upload Sales Template Workbook",
}


uploaded_file = st.file_uploader(
    upload_labels[converter],
    type=["xlsx"]
)


if uploaded_file:

    if st.button("Convert"):

        with st.spinner("Processing workbook..."):

            try:

                 convert_function, output_filename = CONVERTERS[converter]
                 df = convert_function(uploaded_file)

            except Exception as e:
                 st.error(f"❌ {e}")
                 st.stop()

        st.success(f"{len(df):,} rows created.")

        st.dataframe(df.head(100), use_container_width=True)

        output = BytesIO()

        df.to_excel(output, index=False)

        output.seek(0)

        st.download_button(
            label="Download Database",
            data=output,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )