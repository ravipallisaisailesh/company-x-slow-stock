
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(
    page_title="Company X | Slow Stock Review",
    page_icon="📊",
    layout="wide"
)

st.title("Company X — Slow Stock Review")
st.caption("Monthly stock ageing, action classification and customer targeting")

st.info(
    "Upload the original Excel workbook containing Item_Master, "
    "Stock_Aging, Sales_Lines and Customers."
)

uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx"])

REQUIRED = {
    "Item_Master": [
        "SKU", "Item Name", "Category", "Master Pack Qty",
        "Unit Cost (Rs)", "Alternate Item (name)",
        "Alternate Flag", "Status"
    ],
    "Stock_Aging": [
        "SKU", "Qty 0-30 days", "Qty 31-60 days",
        "Qty 61-90 days", "Qty 91-180 days", "Qty 180+ days"
    ],
    "Sales_Lines": [
        "Month", "CustomerID", "CustomerName", "Industry",
        "SKU", "Qty", "Value (Rs)"
    ],
    "Customers": ["CustomerID", "CustomerName", "Industry", "Area"]
}

def clean_sku(value):
    if pd.isna(value):
        return None
    return str(value).strip()

def number(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)

def make_excel(sheets):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()

if uploaded:
    try:
        book = pd.ExcelFile(uploaded)
        missing_sheets = [s for s in REQUIRED if s not in book.sheet_names]

        if missing_sheets:
            st.error("Missing sheets: " + ", ".join(missing_sheets))
            st.stop()

        items = pd.read_excel(book, sheet_name="Item_Master")
        stock = pd.read_excel(book, sheet_name="Stock_Aging")
        sales = pd.read_excel(book, sheet_name="Sales_Lines")
        customers = pd.read_excel(book, sheet_name="Customers")

        problems = []

        # Validate required columns
        for name, required_columns in REQUIRED.items():
            frame = {
                "Item_Master": items,
                "Stock_Aging": stock,
                "Sales_Lines": sales,
                "Customers": customers
            }[name]
            absent = [c for c in required_columns if c not in frame.columns]
            if absent:
                st.error(f"{name} is missing columns: {', '.join(absent)}")
                st.stop()

        # Standardise identifiers and numeric fields
        for frame in [items, stock, sales]:
            frame["SKU"] = frame["SKU"].map(clean_sku)

        customers["CustomerID"] = customers["CustomerID"].map(clean_sku)
        sales["CustomerID"] = sales["CustomerID"].map(clean_sku)

        numeric_stock = [
            "Qty 0-30 days", "Qty 31-60 days", "Qty 61-90 days",
            "Qty 91-180 days", "Qty 180+ days"
        ]
        for col in numeric_stock:
            original = stock[col].copy()
            converted = pd.to_numeric(original, errors="coerce")
            bad = original.notna() & converted.isna()
            if bad.any():
                problems.append(
                    f"Stock_Aging: {int(bad.sum())} non-numeric values in {col}; treated as zero."
                )
            stock[col] = converted.fillna(0)

        for col in ["Master Pack Qty", "Unit Cost (Rs)"]:
            original = items[col].copy()
            converted = pd.to_numeric(original, errors="coerce")
            bad = original.notna() & converted.isna()
            if bad.any():
                problems.append(
                    f"Item_Master: {int(bad.sum())} non-numeric values in {col}."
                )
            items[col] = converted

        for col in ["Qty", "Value (Rs)"]:
            original = sales[col].copy()
            converted = pd.to_numeric(original, errors="coerce")
            bad = original.notna() & converted.isna()
            if bad.any():
                problems.append(
                    f"Sales_Lines: {int(bad.sum())} non-numeric values in {col}; treated as zero."
                )
            sales[col] = converted.fillna(0)

        if items["SKU"].duplicated().any():
            duplicates = items.loc[items["SKU"].duplicated(False), "SKU"].dropna().unique()
            problems.append(
                "Duplicate item master SKUs found: " + ", ".join(map(str, duplicates))
                + ". Review before relying on the results."
            )

        if stock["SKU"].duplicated().any():
            problems.append("Duplicate stock rows found for one or more SKUs; quantities are summed.")

        if items["SKU"].isna().any():
            problems.append("Item master has missing SKUs; those rows cannot be matched reliably.")

        # Resolve SWITCH mappings: alternate item is stored by name
        name_to_sku = {}
        for _, row in items.iterrows():
            if pd.notna(row["Item Name"]):
                name_to_sku[str(row["Item Name"]).strip().casefold()] = row["SKU"]

        switch_map = {}
        for _, row in items.iterrows():
            flag = str(row["Alternate Flag"]).strip().upper() if pd.notna(row["Alternate Flag"]) else ""
            if flag == "SWITCH":
                alternate = row["Alternate Item (name)"]
                if pd.isna(alternate):
                    problems.append(f"{row['SKU']}: SWITCH flag has no alternate item name.")
                    continue
                new_sku = name_to_sku.get(str(alternate).strip().casefold())
                if new_sku is None:
                    problems.append(
                        f"{row['SKU']}: alternate item '{alternate}' was not found in Item_Master."
                    )
                elif new_sku == row["SKU"]:
                    problems.append(f"{row['SKU']}: SWITCH points to the same SKU.")
                else:
                    switch_map[row["SKU"]] = new_sku

        # Resolve multi-step switch chains safely
        def resolve_sku(sku):
            sku = clean_sku(sku)
            seen = set()
            while sku in switch_map and sku not in seen:
                seen.add(sku)
                sku = switch_map[sku]
            if sku in seen:
                problems.append(f"Switch mapping cycle detected near SKU {sku}.")
                return None
            return sku

        items["Final SKU"] = items["SKU"].map(resolve_sku)
        stock["Final SKU"] = stock["SKU"].map(resolve_sku)
        sales["Final SKU"] = sales["SKU"].map(resolve_sku)

        # Warn about EXCLUDE flags; the brief does not define this flag's meaning
        exclude_rows = items[
            items["Alternate Flag"].astype(str).str.strip().str.upper().eq("EXCLUDE")
        ]
        if not exclude_rows.empty:
            problems.append(
                "EXCLUDE flag found for SKU(s): "
                + ", ".join(exclude_rows["SKU"].dropna().astype(str).tolist())
                + ". The case brief does not define this flag; confirm the intended treatment with the data owner."
            )

        # Aggregate stock quantities after SWITCH consolidation
        stock_totals = stock.dropna(subset=["Final SKU"]).groupby("Final SKU")[numeric_stock].sum()
        stock_totals["Total Stock"] = stock_totals[numeric_stock].sum(axis=1)
        stock_totals["Aged Qty"] = stock_totals["Qty 91-180 days"] + stock_totals["Qty 180+ days"]

        # Aggregate sales after SWITCH consolidation
        sales_totals = sales.dropna(subset=["Final SKU"]).groupby("Final SKU").agg(
            Six_Month_Sales_Qty=("Qty", "sum"),
            Six_Month_Sales_Value=("Value (Rs)", "sum")
        )

        # Use the new/alternate SKU's master data for switched products
        master = items.drop_duplicates("SKU").set_index("SKU")
        active_master = items[
            ~items["Status"].astype(str).str.strip().str.upper().str.contains("INACTIVE", na=False)
        ].copy()

        # If the status field has an unexpected value, flag it for review
        allowed_status = {"ACTIVE", "INACTIVE"}
        odd_status = items[
            ~items["Status"].astype(str).str.strip().str.upper().isin(allowed_status)
        ]
        if not odd_status.empty:
            problems.append(
                "Unexpected Status values found: "
                + ", ".join(odd_status["Status"].dropna().astype(str).unique())
                + ". Confirm these values with the data owner."
            )

        # Inactive stock warning, including stock on old switch SKUs
        inactive_skus = set(items.loc[
            items["Status"].astype(str).str.strip().str.upper().str.contains("INACTIVE", na=False),
            "SKU"
        ].dropna())
        inactive_stock = stock[
            stock["SKU"].isin(inactive_skus)
        ].copy()
        if not inactive_stock.empty:
            inactive_stock["Stock Qty"] = inactive_stock[numeric_stock].sum(axis=1)
            with_stock = inactive_stock[inactive_stock["Stock Qty"] > 0]
            if not with_stock.empty:
                problems.append(
                    "Inactive SKU(s) still holding stock: "
                    + ", ".join(with_stock["SKU"].dropna().astype(str).unique())
                    + ". Excluded from active-item analysis; investigate disposal or transfer."
                )

        # Analyse active items under their final SKU
        rows = []
        for final_sku in active_master["Final SKU"].dropna().unique():
            if final_sku not in master.index:
                continue
            new = master.loc[final_sku]
            if isinstance(new, pd.DataFrame):
                new = new.iloc[0]

            # Only the final SKU's master pack, cost, category and name are used
            stock_row = stock_totals.loc[final_sku] if final_sku in stock_totals.index else None
            sales_row = sales_totals.loc[final_sku] if final_sku in sales_totals.index else None

            total_stock = float(stock_row["Total Stock"]) if stock_row is not None else 0
            aged_qty = float(stock_row["Aged Qty"]) if stock_row is not None else 0
            aged_value = aged_qty * (float(new["Unit Cost (Rs)"]) if pd.notna(new["Unit Cost (Rs)"]) else 0)
            sold_qty = float(sales_row["Six_Month_Sales_Qty"]) if sales_row is not None else 0
            sold_value = float(sales_row["Six_Month_Sales_Value"]) if sales_row is not None else 0
            pack = float(new["Master Pack Qty"]) if pd.notna(new["Master Pack Qty"]) else 0
            avg_monthly = sold_qty / 6
            months_clear = total_stock / avg_monthly if avg_monthly > 0 else np.inf

            # Rule R5: apply in this exact order
            if pack > 0 and sold_qty < pack:
                classification = "Red List"
            elif aged_qty <= 0 or months_clear < 3:
                classification = "No Action"
            else:
                classification = "Push"

            rows.append({
                "SKU": final_sku,
                "Item Name": new["Item Name"],
                "Category": new["Category"],
                "Master Pack Qty": pack,
                "Unit Cost (Rs)": new["Unit Cost (Rs)"],
                "Total Stock": total_stock,
                "Aged Qty": aged_qty,
                "Aged Value (Rs)": aged_value,
                "Sales Qty (6 months)": sold_qty,
                "Sales Value (6 months)": sold_value,
                "Average Monthly Sales": avg_monthly,
                "Months to Clear": months_clear,
                "Class": classification
            })

        analysis = pd.DataFrame(rows)

        if analysis.empty:
            st.error("No active items could be analysed. Check item status and SKU mappings.")
            st.stop()

        # Customer purchase values across all six months
        customer_spend = sales.groupby("CustomerID", dropna=False).agg(
            Customer_Spend=("Value (Rs)", "sum")
        ).reset_index()
        customer_info = customers.drop_duplicates("CustomerID")
        customer_spend = customer_spend.merge(customer_info, on="CustomerID", how="left", suffixes=("", "_master"))
        eligible = customer_spend[customer_spend["Customer_Spend"] > 100000].copy()

        # Metadata and buyer history for recommendations
        item_meta = items.drop_duplicates("SKU").set_index("SKU")
        category_by_final = {}
        for _, row in items.iterrows():
            if pd.notna(row["Final SKU"]):
                category_by_final[row["Final SKU"]] = row["Category"]

        buyer_ids = {}
        buyer_industries = {}
        for sku, group in sales.dropna(subset=["Final SKU"]).groupby("Final SKU"):
            buyer_ids[sku] = set(group.loc[group["Qty"] > 0, "CustomerID"].dropna())
            buyer_industries[sku] = set(
                group.loc[
                    (group["Qty"] > 0) & group["Industry"].notna(),
                    "Industry"
                ].astype(str).str.strip()
            ) - {""}

        customer_bought_categories = {}
        for customer_id, group in sales.dropna(subset=["CustomerID", "Final SKU"]).groupby("CustomerID"):
            categories = set()
            for sku in group.loc[group["Qty"] > 0, "Final SKU"]:
                category = category_by_final.get(sku)
                if pd.notna(category):
                    categories.add(str(category).strip())
            customer_bought_categories[customer_id] = categories

        targets_by_sku = {}
        push_skus = set(analysis.loc[analysis["Class"] == "Push", "SKU"])

        for sku in push_skus:
            already_bought = buyer_ids.get(sku, set())
            industries = buyer_industries.get(sku, set())
            category = category_by_final.get(sku)

            target_rows = []
            for _, customer in eligible.iterrows():
                cid = customer["CustomerID"]
                industry = customer.get("Industry")
                bucket = None

                # R6: assign only the first applicable bucket
                if cid in already_bought:
                    bucket = 1
                elif pd.notna(industry) and str(industry).strip() in industries:
                    bucket = 2
                elif pd.notna(category) and str(category).strip() in customer_bought_categories.get(cid, set()):
                    bucket = 3

                if bucket is not None:
                    target_rows.append({
                        "SKU": sku,
                        "Bucket": bucket,
                        "CustomerID": cid,
                        "CustomerName": customer.get("CustomerName"),
                        "Industry": industry,
                        "Area": customer.get("Area"),
                        "6-Month Spend (Rs)": customer["Customer_Spend"]
                    })

            targets_by_sku[sku] = pd.DataFrame(target_rows).sort_values(
                ["Bucket", "6-Month Spend (Rs)"],
                ascending=[True, False]
            ) if target_rows else pd.DataFrame(
                columns=["SKU", "Bucket", "CustomerID", "CustomerName",
                         "Industry", "Area", "6-Month Spend (Rs)"]
            )

        # Dashboard summary
        st.header("1. Executive Summary")
        counts = analysis["Class"].value_counts()
        a, b, c, d = st.columns(4)
        a.metric("Red List Items", int(counts.get("Red List", 0)))
        b.metric("No Action Items", int(counts.get("No Action", 0)))
        c.metric("Push Items", int(counts.get("Push", 0)))
        push_aged_value = analysis.loc[
            analysis["Class"] == "Push", "Aged Value (Rs)"
        ].sum()
        d.metric("Push Aged Value", f"₹{push_aged_value:,.2f}")

        st.subheader("Top 10 Push Items by Aged Value")
        top10 = analysis[analysis["Class"] == "Push"].sort_values(
            "Aged Value (Rs)", ascending=False
        ).head(10)
        st.dataframe(top10, use_container_width=True, hide_index=True)

        # Data quality findings
        st.header("2. Data Quality Checks")
        if problems:
            for problem in dict.fromkeys(problems):
                st.warning(problem)
        else:
            st.success("No issues detected by the current automated checks.")

        st.caption(
            "Review all warnings with the data owner. Automated checks cannot establish "
            "the intended meaning of ambiguous source fields."
        )

        # Filterable item list
        st.header("3. Item Analysis")
        class_options = ["All"] + sorted(analysis["Class"].dropna().unique().tolist())
        category_options = ["All"] + sorted(
            analysis["Category"].dropna().astype(str).unique().tolist()
        )
        f1, f2 = st.columns(2)
        selected_class = f1.selectbox("Filter by class", class_options)
        selected_category = f2.selectbox("Filter by category", category_options)

        filtered = analysis.copy()
        if selected_class != "All":
            filtered = filtered[filtered["Class"] == selected_class]
        if selected_category != "All":
            filtered = filtered[filtered["Category"].astype(str) == selected_category]

        st.dataframe(
            filtered.sort_values("Aged Value (Rs)", ascending=False),
            use_container_width=True,
            hide_index=True
        )

        # Customer targeting
        st.header("4. Customer Call List")
        push_options = sorted(push_skus, key=str)
        if push_options:
            selected_sku = st.selectbox("Choose a Push SKU", push_options)
            selected_item = analysis[analysis["SKU"] == selected_sku].iloc[0]
            st.write(f"**Item:** {selected_item['Item Name']} | **Category:** {selected_item['Category']}")
            st.caption("Targets have total six-month purchases greater than ₹1,00,000.")
            target_frame = targets_by_sku.get(selected_sku, pd.DataFrame())
            if target_frame.empty:
                st.info("No eligible customer targets matched the rules for this item.")
            else:
                st.dataframe(target_frame, use_container_width=True, hide_index=True)
        else:
            st.info("There are no Push items, so no customer call list is available.")

        # Export results
        st.header("5. Export")
        export_sheets = {
            "Item_Analysis": analysis,
            "Top_10_Push": top10,
            "Data_Quality": pd.DataFrame({"Finding": list(dict.fromkeys(problems))}),
        }
        for sku, target_frame in targets_by_sku.items():
            safe_sku = "".join(ch if ch.isalnum() else "_" for ch in str(sku))
            export_sheets[("Targets_" + safe_sku)[:31]] = target_frame

        export_data = make_excel(export_sheets)
        st.download_button(
            "Download full analysis as Excel",
            data=export_data,
            file_name="Company_X_Slow_Stock_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred while analysing the workbook: {e}")
        st.exception(e)
else:
    st.warning("Upload the workbook to run the monthly slow-stock review.")
