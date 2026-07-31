import streamlit as st
from io import BytesIO

from converters.benefits import convert as benefits_convert
from converters.expense import convert as expense_convert
from converters.margin_analysis import convert as margin_analysis_convert
from converters.sales import convert as sales_convert
from converters.windsor_dl import convert as windsor_dl_convert


st.set_page_config(
    page_title="Budget Database Toolkit",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Budget Database Toolkit")
st.caption("Convert budget templates into clean databases.")

st.info(
    """
    1. Select the template type.
    2. Upload the Excel workbook.
    3. Click Convert and download the database.
    """
)

CONVERTERS = {
    "Expense Template": (expense_convert, "Expense Database.xlsx"),
    "Benefits Template": (benefits_convert, "Benefits Database.xlsx"),
    "Sales Template": (sales_convert, "Sales Database.xlsx"),
    "Margin Analysis": (margin_analysis_convert, "Margin Analysis Database.xlsx"),
    "Windor DL": (Windsor_DL_convert, "Direct Labour Database.xlsx"),

}

UPLOAD_LABELS = {
    "Expense Template": "Upload Expense Template Workbook",
    "Benefits Template": "Upload Benefits Template Workbook",
    "Sales Template": "Upload Sales Template Workbook",
    "Margin Analysis": "Upload Margin Analysis Workbook",
    "Windsor DL": "Upload Direct Labour Workbook",
}

converter = st.selectbox("Choose a converter", list(CONVERTERS))
uploaded_file = st.file_uploader(UPLOAD_LABELS[converter], type=["xlsx"])

if uploaded_file and st.button("Convert"):
    try:
        with st.spinner("Processing workbook..."):
            convert_function, output_filename = CONVERTERS[converter]
            dataframe = convert_function(uploaded_file)
    except Exception as error:
        st.error(f"❌ {error}")
        st.stop()

    st.success(f"{len(dataframe):,} rows created.")
    st.dataframe(dataframe.head(100), use_container_width=True)

    output = BytesIO()
    dataframe.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    st.download_button(
        label="Download Database",
        data=output,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
