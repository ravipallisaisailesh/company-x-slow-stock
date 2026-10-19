
import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(
    page_title="Company X Slow-Stock Tool",
    layout="wide"
)

st.title("Company X Slow-Stock Action Tool")
st.write("Upload the Company X Excel workbook.")

uploaded_file = st.file_uploader(
    "Upload Excel workbook",
    type=["xlsx"]
)

if uploaded_file:
    try:
        sheets = pd.read_excel(
            uploaded_file,
            sheet_name=None
        )

        st.success("Workbook loaded successfully!")

        required = [
            "Item_Master",
            "Stock_Aging",
            "Sales_Lines",
            "Customers"
        ]

        missing = [s for s in required if s not in sheets]

        if missing:
            st.error(f"Missing sheets: {missing}")
            st.write("Sheets found:", list(sheets.keys()))

        else:
            items = sheets["Item_Master"].copy()
            stock = sheets["Stock_Aging"].copy()
            sales = sheets["Sales_Lines"].copy()

            st.subheader("Workbook column names")

            st.write("Item_Master:", list(items.columns))
            st.write("Stock_Aging:", list(stock.columns))
            st.write("Sales_Lines:", list(sales.columns))

            st.info(
                "Check the column names above. "
                "The calculation section will be enabled "
                "after the column names are matched."
            )

    except Exception as e:
        st.error(f"Unable to read workbook: {e}")
