
import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Company X Slow-Stock Tool",
    layout="wide"
)

st.title("Company X Slow-Stock Action Tool")
st.write("Upload the Company X Excel workbook to inspect its data.")

uploaded_file = st.file_uploader(
    "Upload your Excel workbook",
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        excel = pd.ExcelFile(uploaded_file)

        st.success("Workbook opened successfully!")

        st.subheader("Sheets found in your workbook")
        st.write(excel.sheet_names)

        for sheet in excel.sheet_names:
            df = pd.read_excel(excel, sheet_name=sheet)

            st.subheader(sheet)
            st.write("Number of rows:", len(df))
            st.write("Number of columns:", len(df.columns))
            st.write("Column names:", list(df.columns))

            st.dataframe(df.head(5), use_container_width=True)

    except Exception as e:
        st.error(f"Could not read the workbook: {e}")
