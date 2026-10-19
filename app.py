
import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Company X - Slow Stock Analysis",
    page_icon="📊",
    layout="wide"
)

st.title("Company X - Slow Stock Analysis")
st.write("Upload the original Excel workbook to analyse stock ageing, sales and customer opportunities.")

uploaded_file = st.file_uploader(
    "Upload the Company X Excel workbook",
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        excel = pd.ExcelFile(uploaded_file)
        required_sheets = [
            "Item_Master",
            "Stock_Aging",
            "Sales_Lines",
            "Customers"
        ]

        missing = [
            sheet for sheet in required_sheets
            if sheet not in excel.sheet_names
        ]

        if missing:
            st.error("Missing required sheets: " + ", ".join(missing))
        else:
            items = pd.read_excel(excel, sheet_name="Item_Master")
            stock = pd.read_excel(excel, sheet_name="Stock_Aging")
            sales = pd.read_excel(excel, sheet_name="Sales_Lines")
            customers = pd.read_excel(excel, sheet_name="Customers")

            st.success("Workbook uploaded successfully!")

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Items", len(items))
            col2.metric("Stock records", len(stock))
            col3.metric("Sales records", len(sales))
            col4.metric("Customers", len(customers))

            st.subheader("Item Master")
            st.dataframe(items, use_container_width=True)

            st.subheader("Stock Ageing")
            st.dataframe(stock, use_container_width=True)

            st.subheader("Sales Lines")
            st.dataframe(sales, use_container_width=True)

            st.subheader("Customers")
            st.dataframe(customers, use_container_width=True)

            st.download_button(
                "Download uploaded workbook",
                data=uploaded_file.getvalue(),
                file_name="Company_X_Uploaded_Workbook.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as error:
        st.error(f"Could not read the workbook: {error}")
else:
    st.info("Upload the original Excel workbook to begin.")

