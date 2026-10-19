
import streamlit as st

st.set_page_config(
    page_title="Company X Slow-Stock Tool",
    layout="wide"
)

st.title("Company X Slow-Stock Action Tool")

st.write(
    "Upload your Excel data to review slow-moving stock "
    "and identify customers to contact."
)

uploaded_file = st.file_uploader(
    "Upload your Excel workbook",
    type=["xlsx"]
)

if uploaded_file is not None:
    st.success("Excel file uploaded successfully!")
    st.write("File name:", uploaded_file.name)
