
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(
    page_title="Company X Slow-Stock Tool",
    layout="wide"
)

st.title("Company X Slow-Stock Action Tool")
st.write("Upload the Company X Excel workbook to identify slow stock and customer actions.")

uploaded_file = st.file_uploader(
    "Upload Excel workbook",
    type=["xlsx"]
)

REQUIRED_SHEETS = [
    "Item_Master", "Stock_Aging", "Sales_Lines", "Customers"
]

ITEM_COLS = [
    "SKU", "Item Name", "Category", "Master Pack Qty",
    "Unit Cost (Rs)", "Alternate Item (name)",
    "Alternate Flag", "Status"
]
STOCK_COLS = [
    "SKU", "Qty 0-30 days", "Qty 31-60 days",
    "Qty 61-90 days", "Qty 91-180 days", "Qty 180+ days"
]
SALES_COLS = [
    "Month", "CustomerID", "CustomerName", "Industry",
    "SKU", "Qty", "Value (Rs)"
]
CUSTOMER_COLS = ["CustomerID", "CustomerName", "Industry", "Area"]

def clean_sku(series):
    return series.astype("string").str.strip()

def resolve_sku(sku, mapping):
    seen = set()
    while sku in mapping and sku not in seen:
        seen.add(sku)
        sku = mapping[sku]
    return sku

if uploaded_file is not None:
    try:
        sheets = pd.read_excel(uploaded_file, sheet_name=None)

        missing = [s for s in REQUIRED_SHEETS if s not in sheets]
        if missing:
            st.error(f"Workbook is missing these sheets: {missing}")
            st.stop()

        items = sheets["Item_Master"].copy()
        stock = sheets["Stock_Aging"].copy()
        sales = sheets["Sales_Lines"].copy()
        customers = sheets["Customers"].copy()

        # Validate exact required columns
        required_columns = {
            "Item_Master": ITEM_COLS,
            "Stock_Aging": STOCK_COLS,
            "Sales_Lines": SALES_COLS,
            "Customers": CUSTOMER_COLS,
        }
        frames = {
            "Item_Master": items,
            "Stock_Aging": stock,
            "Sales_Lines": sales,
            "Customers": customers,
        }

        for sheet_name, required in required_columns.items():
            missing_cols = [
                c for c in required if c not in frames[sheet_name].columns
            ]
            if missing_cols:
                st.error(
                    f"{sheet_name} is missing columns: {missing_cols}"
                )
                st.stop()

        # Clean text and identifiers
        items["SKU"] = clean_sku(items["SKU"])
        stock["SKU"] = clean_sku(stock["SKU"])
        sales["SKU"] = clean_sku(sales["SKU"])
        customers["CustomerID"] = clean_sku(customers["CustomerID"])
        sales["CustomerID"] = clean_sku(sales["CustomerID"])

        for col in ["Item Name", "Category", "Alternate Item (name)",
                    "Alternate Flag", "Status"]:
            items[col] = items[col].astype("string").str.strip()

        items["Alternate Flag"] = items["Alternate Flag"].str.upper()
        items["Status"] = items["Status"].str.upper()

        # Convert numeric fields
        item_numeric = ["Master Pack Qty", "Unit Cost (Rs)"]
        stock_numeric = STOCK_COLS[1:]
        sales_numeric = ["Qty", "Value (Rs)"]

        for col in item_numeric:
            items[col] = pd.to_numeric(items[col], errors="coerce")
        for col in stock_numeric:
            stock[col] = pd.to_numeric(stock[col], errors="coerce").fillna(0)
        for col in sales_numeric:
            sales[col] = pd.to_numeric(sales[col], errors="coerce").fillna(0)

        if items["SKU"].duplicated().any():
            st.error("Item_Master contains duplicate SKU values. Correct these before analysis.")
            st.stop()

        # R1: Map old SWITCH SKUs to the SKU of the named replacement item
        name_to_sku = {
            str(row["Item Name"]).strip().casefold(): row["SKU"]
            for _, row in items.iterrows()
            if pd.notna(row["Item Name"]) and pd.notna(row["SKU"])
        }

        switch_map = {}
        switch_issues = []

        for _, row in items.iterrows():
            if row["Alternate Flag"] == "SWITCH":
                alternate_name = row["Alternate Item (name)"]
                if pd.isna(alternate_name):
                    switch_issues.append(f"Missing alternate name for {row['SKU']}")
                    continue

                new_sku = name_to_sku.get(str(alternate_name).strip().casefold())
                if new_sku is None:
                    switch_issues.append(
                        f"Alternate item not found for {row['SKU']}: {alternate_name}"
                    )
                else:
                    switch_map[row["SKU"]] = new_sku

        items["Final SKU"] = items["SKU"].map(
            lambda sku: resolve_sku(sku, switch_map)
        )
        stock["Final SKU"] = stock["SKU"].map(
            lambda sku: resolve_sku(sku, switch_map)
        )
        sales["Final SKU"] = sales["SKU"].map(
            lambda sku: resolve_sku(sku, switch_map)
        )

        # R2: Exclude inactive items from the action list; separately warn about stock
        inactive_mask = items["Status"].str.contains(
            "INACTIVE", na=False
        )
        inactive_skus = set(items.loc[inactive_mask, "SKU"])

        stock_by_sku = stock.groupby("SKU", dropna=False)[stock_numeric].sum()
        inactive_stock_rows = []
        for sku in inactive_skus:
            if sku in stock_by_sku.index:
                qty = float(stock_by_sku.loc[sku].sum())
                if qty > 0:
                    inactive_stock_rows.append({"SKU": sku, "Total Stock": qty})

        # Use the active replacement SKU's item metadata
        active_items = items.loc[~inactive_mask].copy()
        final_items = active_items.loc[
            active_items["SKU"] == active_items["Final SKU"]
        ].copy()

        # R1: Consolidate stock and sales for old and replacement SKUs
        stock_by_final = stock.groupby("Final SKU", dropna=False)[stock_numeric].sum().reset_index()
        sales_by_final = sales.groupby("Final SKU", dropna=False).agg(
            Total_Sold_Qty=("Qty", "sum"),
            Total_Sales_Value=("Value (Rs)", "sum")
        ).reset_index()

        analysis = final_items.merge(
            stock_by_final, on="Final SKU", how="left"
        ).merge(
            sales_by_final, on="Final SKU", how="left"
        )

        for col in stock_numeric:
            analysis[col] = analysis[col].fillna(0)
        analysis["Total_Sold_Qty"] = analysis["Total_Sold_Qty"].fillna(0)
        analysis["Total_Sales_Value"] = analysis["Total_Sales_Value"].fillna(0)

        # R3 and R4: aged stock, aged value, average monthly sales and clearance time
        analysis["Total Stock"] = analysis[stock_numeric].sum(axis=1)
        analysis["Aged Qty"] = (
            analysis["Qty 91-180 days"] + analysis["Qty 180+ days"]
        )
        analysis["Aged Stock Value (Rs)"] = (
            analysis["Aged Qty"] * analysis["Unit Cost (Rs)"].fillna(0)
        )
        analysis["Avg Monthly Sales Qty"] = analysis["Total_Sold_Qty"] / 6

        analysis["Months to Clear"] = np.where(
            analysis["Avg Monthly Sales Qty"] > 0,
            analysis["Total Stock"] / analysis["Avg Monthly Sales Qty"],
            np.inf
        )

        # R5: Apply classification in the required order
        def classify(row):
            pack_qty = row["Master Pack Qty"]
            if pd.isna(pack_qty):
                return "Review Data"
            if row["Total_Sold_Qty"] < pack_qty:
                return "Red List"
            if row["Aged Qty"] <= 0 or row["Months to Clear"] < 3:
                return "No Action"
            return "Push"

        analysis["Action"] = analysis.apply(classify, axis=1)

        # R6: Prepare customer spend and buying history
        customer_spend = sales.groupby("CustomerID", dropna=False).agg(
            Total_Customer_Spend=("Value (Rs)", "sum")
        ).reset_index()

        customer_info = customers.drop_duplicates("CustomerID").copy()
        eligible = customer_info.merge(
            customer_spend, on="CustomerID", how="left"
        )
        eligible["Total_Customer_Spend"] = eligible[
            "Total_Customer_Spend"
        ].fillna(0)
        eligible = eligible.loc[
            eligible["Total_Customer_Spend"] > 100000
        ].copy()

        # Use consolidated SKU history for buying relationships
        buyers_by_sku = sales.groupby("Final SKU")["CustomerID"].apply(
            lambda s: set(s.dropna())
        ).to_dict()

        industries_by_sku = sales.groupby("Final SKU")["Industry"].apply(
            lambda s: set(
                x for x in s.dropna().astype(str).str.strip()
                if x and x.casefold() != "nan"
            )
        ).to_dict()

        category_by_sku = dict(zip(analysis["Final SKU"], analysis["Category"]))

        targets = []
        push_skus = analysis.loc[
            analysis["Action"] == "Push", "Final SKU"
        ].tolist()

        for sku in push_skus:
            item_buyers = buyers_by_sku.get(sku, set())
            item_industries = industries_by_sku.get(sku, set())
            category = category_by_sku.get(sku)

            for _, cust in eligible.iterrows():
                cid = cust["CustomerID"]
                bucket = None
                reason = None

                if cid in item_buyers:
                    bucket = 1
                    reason = "Previously bought this item or its SWITCH partner"

                elif (
                    pd.notna(cust["Industry"])
                    and str(cust["Industry"]).strip() in item_industries
                    and len(item_industries) > 0
                ):
                    bucket = 2
                    reason = "Same industry as a customer who bought this item"

                elif pd.notna(category):
                    customer_skus = sales.loc[
                        sales["CustomerID"] == cid, "Final SKU"
                    ].dropna().unique()
                    customer_categories = {
                        category_by_sku.get(s) for s in customer_skus
                    }
                    if category in customer_categories:
                        bucket = 3
                        reason = "Bought another item in the same category"

                if bucket is not None:
                    targets.append({
                        "SKU": sku,
                        "CustomerID": cid,
                        "CustomerName": cust["CustomerName"],
                        "Industry": cust["Industry"],
                        "Area": cust["Area"],
                        "Total Customer Spend (Rs)": cust["Total_Customer_Spend"],
                        "Bucket": bucket,
                        "Reason": reason
                    })

        targets_df = pd.DataFrame(targets)

        if not targets_df.empty:
            targets_df = targets_df.sort_values(
                ["SKU", "Bucket", "Total Customer Spend (Rs)"],
                ascending=[True, True, False]
            )

        # Data quality warnings
        problems = []
        if sales["CustomerID"].isna().any():
            problems.append("Sales_Lines contains missing CustomerID values.")
        if sales["Final SKU"].isna().any():
            problems.append("Sales_Lines contains missing SKU values.")
        if (stock[stock_numeric] < 0).any().any():
            problems.append("Stock_Aging contains negative stock quantities.")
        if items["Master Pack Qty"].isna().any():
            problems.append("Some items have missing or invalid Master Pack Qty.")
        if switch_issues:
            problems.extend(switch_issues)

        st.success("Workbook loaded and analysis completed.")

        # Dashboard
        st.subheader("Executive Summary")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Active Final SKUs", len(analysis))
        c2.metric("Red List", int((analysis["Action"] == "Red List").sum()))
        c3.metric("No Action", int((analysis["Action"] == "No Action").sum()))
        c4.metric("Push", int((analysis["Action"] == "Push").sum()))

        push_value = analysis.loc[
            analysis["Action"] == "Push", "Aged Stock Value (Rs)"
        ].sum()
        st.metric("Aged Stock Value in Push Items", f"₹{push_value:,.2f}")

        st.subheader("Slow-Stock Action List")
        st.caption("Months to Clear shown as infinity means there are no recorded sales.")
        display_cols = [
            "Final SKU", "Item Name", "Category", "Master Pack Qty",
            "Total Stock", "Total_Sold_Qty", "Aged Qty",
            "Aged Stock Value (Rs)", "Avg Monthly Sales Qty",
            "Months to Clear", "Action"
        ]
        st.dataframe(
            analysis[display_cols].rename(columns={"Final SKU": "SKU"}),
            use_container_width=True
        )

        st.subheader("Customer Recommendations for Push Items")
        if targets_df.empty:
            st.info("No eligible customer recommendations were generated.")
        else:
            st.dataframe(targets_df, use_container_width=True)

        st.subheader("SWITCH Mapping")
        switch_df = pd.DataFrame(
            list(switch_map.items()),
            columns=["Old SKU", "New SKU"]
        )
        st.dataframe(switch_df, use_container_width=True)

        st.subheader("Inactive Items with Stock")
        st.dataframe(
            pd.DataFrame(inactive_stock_rows, columns=["SKU", "Total Stock"]),
            use_container_width=True
        )

        st.subheader("Data Quality Warnings")
        if problems:
            for problem in problems:
                st.warning(problem)
        else:
            st.success("No issues detected by the basic data-quality checks.")

        # Export results to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            analysis.to_excel(writer, sheet_name="Item_List", index=False)
            targets_df.to_excel(writer, sheet_name="Customer_Targets", index=False)
            switch_df.to_excel(writer, sheet_name="Switch_Mapping", index=False)
            pd.DataFrame(inactive_stock_rows).to_excel(
                writer, sheet_name="Inactive_Stock", index=False
            )
            pd.DataFrame({"Warning": problems}).to_excel(
                writer, sheet_name="Data_Problems", index=False
            )

        st.download_button(
            "Download Analysis Excel",
            data=output.getvalue(),
            file_name="Company_X_Slow_Stock_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Analysis failed: {e}")
        st.exception(e)
