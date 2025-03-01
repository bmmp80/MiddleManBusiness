import os
import sqlite3
import pandas as pd
import numpy as np
import sys

def run_all():
    # Define file paths
    # db_file = os.path.join(os.path.dirname(__file__), "business_tracker.db")
    # report_file = os.path.join(os.path.dirname(__file__), "full financial report.csv")

    if getattr(sys, 'frozen', False):
        # If the application is run as a bundled executable, get the directory of the executable.
        base_dir = os.path.dirname(sys.executable)
    else:
        # If it's run as a script, use the directory of the script.
        base_dir = os.path.dirname(__file__)

    db_file = os.path.join(base_dir, "business_tracker.db")
    report_file = os.path.join(base_dir, "full financial report.xlsx")

    # ✅ Step 1: Delete existing report if it exists
    if os.path.exists(report_file):
        os.remove(report_file)
        print("🚨 Previous financial report deleted. Generating a new one.")

    # Ensure the database exists
    if not os.path.exists(db_file):
        print("Database does not exist. Creating schema...")
        # Adjust the path below if needed
        os.system("python C:/Users/brand/PycharmProjects/MiddleManBusiness/Attempt7/CreateDatabase.py")
    else:
        print("Database exists. Skipping schema creation.")

    # ✅ Step 2: Connect to Database and Retrieve Offer Data
    conn = sqlite3.connect(db_file)

    query_offers = """
    SELECT 
        o.id AS Offer_ID,
        CASE WHEN o.purchase_complete = 1 THEN 'true'
             ELSE 'false'
        END AS Purchase_Complete,
        c.name AS SellerName,
        c.site AS Supplier_Email,
        o.expected_receive_date AS Expected_Receive_Date,
        (SELECT MAX(sell_date) FROM CustomerSale WHERE offer_id = o.id) AS Expected_Flip_Date,
        (
            SELECT COALESCE(SUM(op.quantity_purchased * op.purchase_price_per_unit), 0)
            FROM Offer_Product op 
            WHERE op.offer_id = o.id
        ) AS Investment,
        (
            SELECT COUNT(*)
            FROM CustomerSale cs
            WHERE cs.offer_id = o.id
        ) AS NumberOfSales
    FROM Offer o
    JOIN Contact c ON o.contact_id = c.id
    ORDER BY o.id;
    """

    query_products = """
    SELECT 
        o.id AS Offer_ID,
        p.name AS ProductName,
        op.quantity_purchased AS Purchase_Quantity,
        op.purchase_price_per_unit AS Purchase_Cost_Per_Unit,
        op.unit AS Product_Unit,
        (op.quantity_purchased * op.purchase_price_per_unit) AS TotalProductCost
    FROM Offer o
    JOIN Offer_Product op ON o.id = op.offer_id
    JOIN Product p ON op.product_id = p.id
    ORDER BY o.id, p.name;
    """

    query_sales = """
    SELECT 
        o.id AS Offer_ID,
        CASE WHEN cs.sale_complete = 1 THEN 'true'
             ELSE 'false'
        END AS sale_complete,
        cs.customer AS CustomerName,
        p.name AS Sold_ProductName,
        csp.quantity_sold AS Quantity_Sold,
        csp.sell_price_per_unit AS SellPricePerUnit,
        csp.unit AS Sale_Unit,
        (csp.quantity_sold * csp.sell_price_per_unit) AS TotalSalePrice,
        cs.sell_date AS Sell_Date
    FROM Offer o
    JOIN CustomerSale cs ON o.id = cs.offer_id
    JOIN CustomerSale_Product csp ON cs.id = csp.customer_sale_id
    JOIN Product p ON csp.product_id = p.id
    ORDER BY o.id, cs.customer;
    """

    # -------------------- 3. READ DATA INTO DATAFRAMES --------------------
    df_offers = pd.read_sql_query(query_offers, conn)
    df_products = pd.read_sql_query(query_products, conn)
    df_sales = pd.read_sql_query(query_sales, conn)
    conn.close()

    # -------------------- 4. UNIT CONVERSION FUNCTIONS --------------------
    def convert_units(value, from_unit, to_unit):
        """Converts units for calculations if necessary."""
        conversion_rates = {
            ("oz", "g"): 28.3495,
            ("g", "oz"): 1 / 28.3495,
            ("kg", "g"): 1000,
            ("g", "kg"): 0.001,
            ("lb", "kg"): 0.453592,
            ("kg", "lb"): 2.20462,
        }
        if from_unit == to_unit:
            return value
        return value * conversion_rates.get((from_unit, to_unit), 1)

    def get_adjusted_quantity(row):
        offer_id = row["Offer_ID"]
        quantity = row["Quantity_Sold"]
        sale_unit = row["Sale_Unit"]
        # Filter df_products to get the matching product row
        matching_products = df_products[df_products["Offer_ID"] == offer_id]
        # Make sure we have at least one match
        if not matching_products.empty:
            product_unit = matching_products.iloc[0]["Product_Unit"]
        else:
            # If there is no matching product, assume no conversion is needed.
            product_unit = sale_unit
        return convert_units(quantity, sale_unit, product_unit)

    df_sales["Adjusted_Quantity"] = df_sales.apply(get_adjusted_quantity, axis=1)

    # -------------------- 5. CALCULATIONS (Profit, etc.) --------------------
    # Compute total sold for each offer
    total_offer_fully_sold = df_sales.groupby("Offer_ID")["TotalSalePrice"].sum().reset_index()
    total_offer_fully_sold.rename(columns={"TotalSalePrice": "TotalOfferFullySold"}, inplace=True)

    # Merge into df_offers
    df_offers = df_offers.merge(total_offer_fully_sold, on="Offer_ID", how="left")
    df_offers["TotalOfferFullySold"] = df_offers["TotalOfferFullySold"].fillna(0)

    # Merge whether all sales are complete
    offer_sales_complete = df_sales.groupby("Offer_ID")["sale_complete"].all().reset_index()
    offer_sales_complete.rename(columns={"sale_complete": "all_sales_complete"}, inplace=True)
    df_offers = df_offers.merge(offer_sales_complete, on="Offer_ID", how="left")

    # Replace missing values with False
    df_offers["all_sales_complete"] = pd.array(
        np.where(df_offers["all_sales_complete"].isna(), False, df_offers["all_sales_complete"]),
        dtype="boolean"
    )

    # Calculate profit only if all sales are complete; otherwise profit is 0.
    df_offers["Profit"] = df_offers.apply(
        lambda row: (row["TotalOfferFullySold"] - row["Investment"]) if row["all_sales_complete"] else 0,
        axis=1
    )

    # -------------------- 6. BUILD CASCADING DATA --------------------
    cascading_data = []

    offer_columns = [
        "Offer_ID",
        "Purchase_Complete",
        "SellerName",
        "Supplier_Email",
        "Expected_Receive_Date",
        "Expected_Flip_Date",
        "Investment",
        "NumberOfSales",
        "TotalOfferFullySold",
        "Profit",
    ]

    for offer_id in df_offers["Offer_ID"].unique():
        # This returns a 1-row dataframe for the given offer_id
        offer_subset = df_offers[df_offers["Offer_ID"] == offer_id]
        if offer_subset.empty:
            continue

        # We only need the first row from df_offers for each offer
        offer_row = offer_subset.iloc[0]

        # Grab all products for this offer
        product_rows = df_products[df_products["Offer_ID"] == offer_id]
        # Grab all sales for this offer
        sale_rows = df_sales[df_sales["Offer_ID"] == offer_id]

        # The maximum number of lines we need is whichever is bigger: # of product rows or # of sale rows
        max_rows = max(len(product_rows), len(sale_rows))

        for i in range(max_rows):
            # Build the offer portion of the row
            if i == 0:
                offer_data = [offer_row[col] for col in offer_columns]
            else:
                # For subsequent lines, keep the offer columns blank
                offer_data = [""] * len(offer_columns)

            # Build the product portion of the row (5 columns)
            if i < len(product_rows):
                pr = product_rows.iloc[i]
                product_data = [
                    pr["ProductName"],
                    pr["Purchase_Quantity"],
                    pr["Product_Unit"],
                    pr["Purchase_Cost_Per_Unit"],
                    pr["TotalProductCost"],
                ]
            else:
                product_data = ["", "", "", "", ""]

            # Build the sale portion of the row (8 columns)
            if i < len(sale_rows):
                sr = sale_rows.iloc[i]
                sale_data = [
                    sr["sale_complete"],
                    sr["CustomerName"],
                    sr["Sold_ProductName"],
                    sr["Quantity_Sold"],
                    sr["Sale_Unit"],
                    sr["SellPricePerUnit"],
                    sr["TotalSalePrice"],
                    sr["Sell_Date"],
                ]
            else:
                sale_data = ["", "", "", "", "", "", "", ""]

            # Combine everything
            row_data = offer_data + product_data + sale_data
            cascading_data.append(row_data)

    # -------------------- 7. CREATE FINAL DATAFRAME --------------------
    final_columns = [
        # Offer (10 columns)
        "Offer_ID", "Purchase_Complete", "SellerName", "Supplier_Email", "Expected_Receive_Date",
        "Expected_Flip_Date", "Investment", "NumberOfSales", "TotalOfferFullySold", "Profit",
        # Product (5 columns)
        "ProductName", "Purchase_Quantity", "Product_Unit", "Purchase_Cost_Per_Unit", "TotalProductCost",
        # Sale (8 columns)
        "Sale_Complete", "CustomerName", "Sold_ProductName", "Quantity_Sold", "Sale_Unit",
        "SellPricePerUnit", "TotalSalePrice", "Sell_Date"
    ]

    df_cascading = pd.DataFrame(cascading_data, columns=final_columns)

    # -------------------- 8. EXPORT TO EXCEL (or CSV) --------------------
    df_cascading.to_excel(report_file, index=False)
    print(f"✅ Financial report updated and saved at: {report_file}")

run_all()
