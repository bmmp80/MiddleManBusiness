# # import os
# # # import sqlite3
# # # import pandas as pd
# # # import numpy as np
# # # import sys
# # #
# # # def run_all():
# # #     # Define file paths
# # #     # db_file = os.path.join(os.path.dirname(__file__), "business_tracker.db")
# # #     # report_file = os.path.join(os.path.dirname(__file__), "full financial report.csv")
# # #
# # #     if getattr(sys, 'frozen', False):
# # #         # If the application is run as a bundled executable, get the directory of the executable.
# # #         base_dir = os.path.dirname(sys.executable)
# # #     else:
# # #         # If it's run as a script, use the directory of the script.
# # #         base_dir = os.path.dirname(__file__)
# # #
# # #     db_file = os.path.join(base_dir, "business_tracker.db")
# # #     report_file = os.path.join(base_dir, "full financial report.xlsx")
# # #
# # #     # ✅ Step 1: Delete existing report if it exists
# # #     if os.path.exists(report_file):
# # #         os.remove(report_file)
# # #         print("🚨 Previous financial report deleted. Generating a new one.")
# # #
# # #     # Ensure the database exists
# # #     if not os.path.exists(db_file):
# # #         print("Database does not exist. Creating schema...")
# # #         # Adjust the path below if needed
# # #         os.system("python C:/Users/brand/PycharmProjects/MiddleManBusiness/Attempt7/CreateDatabase.py")
# # #     else:
# # #         print("Database exists. Skipping schema creation.")
# # #
# # #     # ✅ Step 2: Connect to Database and Retrieve Offer Data
# # #     conn = sqlite3.connect(db_file)
# # #
# # #     query_offers = """
# # #     SELECT
# # #         o.id AS Offer_ID,
# # #         CASE WHEN o.purchase_complete = 1 THEN 'true'
# # #              ELSE 'false'
# # #         END AS Purchase_Complete,
# # #         c.name AS SellerName,
# # #         c.site AS Supplier_Email,
# # #         o.expected_receive_date AS Expected_Receive_Date,
# # #         (SELECT MAX(sell_date) FROM CustomerSale WHERE offer_id = o.id) AS Expected_Flip_Date,
# # #         (
# # #             SELECT COALESCE(SUM(op.quantity_purchased * op.purchase_price_per_unit), 0)
# # #             FROM Offer_Product op
# # #             WHERE op.offer_id = o.id
# # #         ) AS Investment,
# # #         (
# # #             SELECT COUNT(*)
# # #             FROM CustomerSale cs
# # #             WHERE cs.offer_id = o.id
# # #         ) AS NumberOfSales
# # #     FROM Offer o
# # #     JOIN Contact c ON o.contact_id = c.id
# # #     ORDER BY o.id;
# # #     """
# # #
# # #     query_products = """
# # #     SELECT
# # #         o.id AS Offer_ID,
# # #         p.name AS ProductName,
# # #         op.quantity_purchased AS Purchase_Quantity,
# # #         op.purchase_price_per_unit AS Purchase_Cost_Per_Unit,
# # #         op.unit AS Product_Unit,
# # #         (op.quantity_purchased * op.purchase_price_per_unit) AS TotalProductCost
# # #     FROM Offer o
# # #     JOIN Offer_Product op ON o.id = op.offer_id
# # #     JOIN Product p ON op.product_id = p.id
# # #     ORDER BY o.id, p.name;
# # #     """
# # #
# # #     query_sales = """
# # #     SELECT
# # #         o.id AS Offer_ID,
# # #         CASE WHEN cs.sale_complete = 1 THEN 'true'
# # #              ELSE 'false'
# # #         END AS sale_complete,
# # #         cs.customer AS CustomerName,
# # #         p.name AS Sold_ProductName,
# # #         csp.quantity_sold AS Quantity_Sold,
# # #         csp.sell_price_per_unit AS SellPricePerUnit,
# # #         csp.unit AS Sale_Unit,
# # #         (csp.quantity_sold * csp.sell_price_per_unit) AS TotalSalePrice,
# # #         cs.sell_date AS Sell_Date
# # #     FROM Offer o
# # #     JOIN CustomerSale cs ON o.id = cs.offer_id
# # #     JOIN CustomerSale_Product csp ON cs.id = csp.customer_sale_id
# # #     JOIN Product p ON csp.product_id = p.id
# # #     ORDER BY o.id, cs.customer;
# # #     """
# # #
# # #     # -------------------- 3. READ DATA INTO DATAFRAMES --------------------
# # #     df_offers = pd.read_sql_query(query_offers, conn)
# # #     df_products = pd.read_sql_query(query_products, conn)
# # #     df_sales = pd.read_sql_query(query_sales, conn)
# # #     conn.close()
# # #
# # #     # -------------------- 4. UNIT CONVERSION FUNCTIONS --------------------
# # #     def convert_units(value, from_unit, to_unit):
# # #         """Converts units for calculations if necessary."""
# # #         conversion_rates = {
# # #             ("oz", "g"): 28.3495,
# # #             ("g", "oz"): 1 / 28.3495,
# # #             ("kg", "g"): 1000,
# # #             ("g", "kg"): 0.001,
# # #             ("lb", "kg"): 0.453592,
# # #             ("kg", "lb"): 2.20462,
# # #         }
# # #         if from_unit == to_unit:
# # #             return value
# # #         return value * conversion_rates.get((from_unit, to_unit), 1)
# # #
# # #     def get_adjusted_quantity(row):
# # #         offer_id = row["Offer_ID"]
# # #         quantity = row["Quantity_Sold"]
# # #         sale_unit = row["Sale_Unit"]
# # #         # Filter df_products to get the matching product row
# # #         matching_products = df_products[df_products["Offer_ID"] == offer_id]
# # #         # Make sure we have at least one match
# # #         if not matching_products.empty:
# # #             product_unit = matching_products.iloc[0]["Product_Unit"]
# # #         else:
# # #             # If there is no matching product, assume no conversion is needed.
# # #             product_unit = sale_unit
# # #         return convert_units(quantity, sale_unit, product_unit)
# # #
# # #     df_sales["Adjusted_Quantity"] = df_sales.apply(get_adjusted_quantity, axis=1)
# # #
# # #     # -------------------- 5. CALCULATIONS (Profit, etc.) --------------------
# # #     # Compute total sold for each offer
# # #     total_offer_fully_sold = df_sales.groupby("Offer_ID")["TotalSalePrice"].sum().reset_index()
# # #     total_offer_fully_sold.rename(columns={"TotalSalePrice": "TotalOfferFullySold"}, inplace=True)
# # #
# # #     # Merge into df_offers
# # #     df_offers = df_offers.merge(total_offer_fully_sold, on="Offer_ID", how="left")
# # #     df_offers["TotalOfferFullySold"] = df_offers["TotalOfferFullySold"].fillna(0)
# # #
# # #     # Merge whether all sales are complete
# # #     offer_sales_complete = df_sales.groupby("Offer_ID")["sale_complete"].all().reset_index()
# # #     offer_sales_complete.rename(columns={"sale_complete": "all_sales_complete"}, inplace=True)
# # #     df_offers = df_offers.merge(offer_sales_complete, on="Offer_ID", how="left")
# # #
# # #     # Replace missing values with False
# # #     df_offers["all_sales_complete"] = pd.array(
# # #         np.where(df_offers["all_sales_complete"].isna(), False, df_offers["all_sales_complete"]),
# # #         dtype="boolean"
# # #     )
# # #
# # #     # Calculate profit only if all sales are complete; otherwise profit is 0.
# # #     df_offers["Profit"] = df_offers.apply(
# # #         lambda row: (row["TotalOfferFullySold"] - row["Investment"]) if row["all_sales_complete"] else 0,
# # #         axis=1
# # #     )
# # #
# # #     # -------------------- 6. BUILD CASCADING DATA --------------------
# # #     cascading_data = []
# # #
# # #     offer_columns = [
# # #         "Offer_ID",
# # #         "Purchase_Complete",
# # #         "SellerName",
# # #         "Supplier_Email",
# # #         "Expected_Receive_Date",
# # #         "Expected_Flip_Date",
# # #         "Investment",
# # #         "NumberOfSales",
# # #         "TotalOfferFullySold",
# # #         "Profit",
# # #     ]
# # #
# # #     for offer_id in df_offers["Offer_ID"].unique():
# # #         # This returns a 1-row dataframe for the given offer_id
# # #         offer_subset = df_offers[df_offers["Offer_ID"] == offer_id]
# # #         if offer_subset.empty:
# # #             continue
# # #
# # #         # We only need the first row from df_offers for each offer
# # #         offer_row = offer_subset.iloc[0]
# # #
# # #         # Grab all products for this offer
# # #         product_rows = df_products[df_products["Offer_ID"] == offer_id]
# # #         # Grab all sales for this offer
# # #         sale_rows = df_sales[df_sales["Offer_ID"] == offer_id]
# # #
# # #         # The maximum number of lines we need is whichever is bigger: # of product rows or # of sale rows
# # #         max_rows = max(len(product_rows), len(sale_rows))
# # #
# # #         for i in range(max_rows):
# # #             # Build the offer portion of the row
# # #             if i == 0:
# # #                 offer_data = [offer_row[col] for col in offer_columns]
# # #             else:
# # #                 # For subsequent lines, keep the offer columns blank
# # #                 offer_data = [""] * len(offer_columns)
# # #
# # #             # Build the product portion of the row (5 columns)
# # #             if i < len(product_rows):
# # #                 pr = product_rows.iloc[i]
# # #                 product_data = [
# # #                     pr["ProductName"],
# # #                     pr["Purchase_Quantity"],
# # #                     pr["Product_Unit"],
# # #                     pr["Purchase_Cost_Per_Unit"],
# # #                     pr["TotalProductCost"],
# # #                 ]
# # #             else:
# # #                 product_data = ["", "", "", "", ""]
# # #
# # #             # Build the sale portion of the row (8 columns)
# # #             if i < len(sale_rows):
# # #                 sr = sale_rows.iloc[i]
# # #                 sale_data = [
# # #                     sr["sale_complete"],
# # #                     sr["CustomerName"],
# # #                     sr["Sold_ProductName"],
# # #                     sr["Quantity_Sold"],
# # #                     sr["Sale_Unit"],
# # #                     sr["SellPricePerUnit"],
# # #                     sr["TotalSalePrice"],
# # #                     sr["Sell_Date"],
# # #                 ]
# # #             else:
# # #                 sale_data = ["", "", "", "", "", "", "", ""]
# # #
# # #             # Combine everything
# # #             row_data = offer_data + product_data + sale_data
# # #             cascading_data.append(row_data)
# # #
# # #     # -------------------- 7. CREATE FINAL DATAFRAME --------------------
# # #     final_columns = [
# # #         # Offer (10 columns)
# # #         "Offer_ID", "Purchase_Complete", "SellerName", "Supplier_Email", "Expected_Receive_Date",
# # #         "Expected_Flip_Date", "Investment", "NumberOfSales", "TotalOfferFullySold", "Profit",
# # #         # Product (5 columns)
# # #         "ProductName", "Purchase_Quantity", "Product_Unit", "Purchase_Cost_Per_Unit", "TotalProductCost",
# # #         # Sale (8 columns)
# # #         "Sale_Complete", "CustomerName", "Sold_ProductName", "Quantity_Sold", "Sale_Unit",
# # #         "SellPricePerUnit", "TotalSalePrice", "Sell_Date"
# # #     ]
# # #
# # #     df_cascading = pd.DataFrame(cascading_data, columns=final_columns)
# # #
# # # #     # -------------------- 8. EXPORT TO EXCEL (or CSV) --------------------
# # # #     df_cascading.to_excel(report_file, index=False)
# # # #     print(f"✅ Financial report updated and saved at: {report_file}")
# # # #
# # # # run_all()
# # # # import pandas as pd
# # # # print("Pandas imported successfully!")
# # # # import time
# # # # import random
# # # # import re
# # # #
# # # # # ===================== CONFIGURATION =====================
# # # # if getattr(sys, 'frozen', False):
# # # #     base_dir = os.path.dirname(sys.executable)
# # # # else:
# # # #     base_dir = os.path.dirname(__file__)
# # # # print(base_dir)
# # # # DB_FILE = os.path.join(base_dir, "business_tracker.db")
# # #
# # # import pandas as pd
# # # print("Pandas imported successfully!")
# # #
# # # import sys
# # # import UpdateDatabaseWithCalculatedFields
# # # import sqlite3
# # # import os
# # # import pandas as pd
# # # import xlwings as xw
# # # import dataaddagain
# # # # ===================== CONFIGURATION =====================
# # # if getattr(sys, 'frozen', False):
# # #     base_dir = os.path.dirname(sys.executable)
# # # else:
# # #     base_dir = os.path.dirname(__file__)
# # # print(base_dir)
# # # DB_FILE = os.path.join(base_dir, "business_tracker.db")
# # # REPORT_FILE = os.path.join(base_dir, "full_analytics_report.xlsx")
# # #
# # #
# # # def execute_query(query, params=(), fetch=False):
# # #     conn = sqlite3.connect(DB_FILE)
# # #     cursor = conn.cursor()
# # #     cursor.execute(query, params)
# # #     data = cursor.fetchall() if fetch else None
# # #     conn.commit()
# # #     conn.close()
# # #     return data
# # #
# # #
# # #
# # # def create_pivot_with_xlwings(excel_file, raw_sheet_name, pivot_sheet_name):
# # #     """
# # #     Opens the Excel workbook and creates a pivot table on a new sheet named pivot_sheet_name.
# # #     It uses data from the sheet specified by raw_sheet_name.
# # #     If raw_sheet_name is not found, it falls back to "Sheet1".
# # #     """
# # #     app = xw.App(visible=False)
# # #     wb = app.books.open(excel_file)
# # #
# # #     # Print available sheet names for debugging
# # #     available_sheets = [s.name for s in wb.sheets]
# # #     print("Available sheets:", available_sheets)
# # #
# # #     # Use raw_sheet_name if it exists; otherwise, fall back to "Sheet1"
# # #     if raw_sheet_name in available_sheets:
# # #         raw_sheet = wb.sheets[raw_sheet_name]
# # #     elif "Sheet1" in available_sheets:
# # #         print(f"Sheet '{raw_sheet_name}' not found. Falling back to 'Sheet1'.")
# # #         raw_sheet = wb.sheets["Sheet1"]
# # #     else:
# # #         print(f"Error: Neither '{raw_sheet_name}' nor 'Sheet1' were found!")
# # #         wb.close()
# # #         app.quit()
# # #         return
# # #
# # #     # Create (or clear) the pivot sheet
# # #     if pivot_sheet_name in available_sheets:
# # #         pivot_sheet = wb.sheets[pivot_sheet_name]
# # #         pivot_sheet.clear()
# # #     else:
# # #         pivot_sheet = wb.sheets.add(pivot_sheet_name)
# # #
# # #     source_range = raw_sheet.range("A1").expand()
# # #
# # #     pivot_table = pivot_sheet.api.PivotTableWizard(
# # #         SourceType=1,  # xlDatabase
# # #         SourceData=source_range.api,
# # #         TableDestination=pivot_sheet.range("A3").api,
# # #         TableName="PivotTable1"
# # #     )
# # #
# # #     # Configure pivot fields (example)
# # #     pivot_table.PivotFields("CustomerName").Orientation = 1  # Row field
# # #     pivot_table.PivotFields("Sell_Date").Orientation = 1     # Row field
# # #     pivot_table.PivotFields("Sold_ProductName").Orientation = 1  # Row field
# # #
# # #     data_field = pivot_table.PivotFields("TotalSalePrice")
# # #     data_field.Orientation = 4  # Data field
# # #     data_field.Function = -4157  # xlSum
# # #     data_field.Name = "Sum of TotalSalePrice"
# # #
# # #     wb.save()
# # #     wb.close()
# # #     app.quit()
# # #     print(f"Pivot table created on sheet '{pivot_sheet_name}' in '{excel_file}'.")
# # #
# # #
# # #
# # # def calculate_and_update_inventory_stats(sheet_name):
# # #     """
# # #     Calculates inventory statistics for every product using precomputed data from
# # #     Offer_Product and CustomerSale_Product, stores the results in a DataFrame,
# # #     and exports that DataFrame to an Excel file (deleting any existing file first).
# # #
# # #     The query returns 15 columns:
# # #       1. product_id
# # #       2. total_purchased (SUM of quantity_purchased)
# # #       3. remaining_inventory (total purchased - total sold)
# # #       4. total_purchase_cost (SUM(quantity_purchased * purchase_price_per_unit))
# # #       5. total_sold (SUM of quantity_sold)
# # #       6. total_sale_revenue (SUM(quantity_sold * sell_price_per_unit))
# # #       7. profit_all_time (sale revenue minus purchase cost)
# # #       8. avg_flip_amount (AVG of sell_price_per_unit)
# # #       9. min_purchase_price (MIN of purchase_price_per_unit)
# # #      10. best_purchase_date (offer_end_date of the offer with the lowest purchase price)
# # #      11. best_purchase_contact (Contact name for that best purchase)
# # #      12. best_customer_id (the customer/contact who purchased the most quantity)
# # #      13. best_customer_total_purchased (the corresponding total quantity)
# # #      14. best_sale_amount (MAX of sell_price_per_unit)
# # #      15. best_sale_date (sell_date corresponding to the best sale amount)
# # #     """
# # #
# # #
# # #     # Get all products (the keys are product IDs as strings)
# # #     prod_dict = UpdateDatabaseWithCalculatedFields.get_products()
# # #     results = []
# # #
# # #     # This query uses 17 placeholders (each subquery gets the same product_id)
# # #     query_inventory_statistics = """
# # #         SELECT
# # #             ? AS product_id,
# # #             (SELECT COALESCE(SUM(quantity_purchased), 0)
# # #              FROM Offer_Product WHERE product_id = ?),
# # #             ((SELECT COALESCE(SUM(quantity_purchased), 0)
# # #               FROM Offer_Product WHERE product_id = ?)
# # #              - (SELECT COALESCE(SUM(quantity_sold), 0)
# # #                 FROM CustomerSale_Product WHERE product_id = ?)),
# # #             (SELECT COALESCE(SUM(quantity_purchased * purchase_price_per_unit), 0)
# # #              FROM Offer_Product WHERE product_id = ?),
# # #             (SELECT COALESCE(SUM(quantity_sold), 0)
# # #              FROM CustomerSale_Product WHERE product_id = ?),
# # #             (SELECT COALESCE(SUM(quantity_sold * sell_price_per_unit), 0)
# # #              FROM CustomerSale_Product WHERE product_id = ?),
# # #             ((SELECT COALESCE(SUM(quantity_sold * sell_price_per_unit), 0)
# # #               FROM CustomerSale_Product WHERE product_id = ?)
# # #              - (SELECT COALESCE(SUM(quantity_purchased * purchase_price_per_unit), 0)
# # #                 FROM Offer_Product WHERE product_id = ?)),
# # #             (SELECT AVG(sell_price_per_unit)
# # #              FROM CustomerSale_Product WHERE product_id = ?),
# # #             (SELECT MIN(purchase_price_per_unit)
# # #              FROM Offer_Product WHERE product_id = ?),
# # #             (SELECT o.offer_end_date
# # #              FROM Offer o JOIN Offer_Product op ON o.id = op.offer_id
# # #              WHERE op.product_id = ? ORDER BY op.purchase_price_per_unit ASC LIMIT 1),
# # #             (SELECT c.name
# # #              FROM Contact c JOIN Offer o ON o.contact_id = c.id JOIN Offer_Product op ON o.id = op.offer_id
# # #              WHERE op.product_id = ? ORDER BY op.purchase_price_per_unit ASC LIMIT 1),
# # #             (SELECT customer_id
# # #              FROM (
# # #                  SELECT cs.contact_id AS customer_id, SUM(csp.quantity_sold) AS total_qty
# # #                  FROM CustomerSale_Product csp
# # #                  JOIN CustomerSale cs ON cs.id = csp.customer_sale_id
# # #                  WHERE csp.product_id = ?
# # #                  GROUP BY cs.contact_id
# # #                  ORDER BY total_qty DESC LIMIT 1
# # #              )),
# # #             (SELECT total_qty
# # #              FROM (
# # #                  SELECT cs.contact_id AS customer_id, SUM(csp.quantity_sold) AS total_qty
# # #                  FROM CustomerSale_Product csp
# # #                  JOIN CustomerSale cs ON cs.id = csp.customer_sale_id
# # #                  WHERE csp.product_id = ?
# # #                  GROUP BY cs.contact_id
# # #                  ORDER BY total_qty DESC LIMIT 1
# # #              )),
# # #             (SELECT MAX(sell_price_per_unit)
# # #              FROM CustomerSale_Product WHERE product_id = ?),
# # #             (SELECT cs.sell_date
# # #              FROM CustomerSale cs JOIN CustomerSale_Product csp ON cs.id = csp.customer_sale_id
# # #              WHERE csp.product_id = ? ORDER BY csp.sell_price_per_unit DESC LIMIT 1)
# # #         """
# # #     # For each product, supply the product_id 17 times
# # #     for prod_id in prod_dict.keys():
# # #         params = tuple([prod_id] * 17)
# # #         row = execute_query(query_inventory_statistics, params, fetch=True)
# # #         if row:
# # #             results.append(row[0])
# # #
# # #     # Define column names (15 columns)
# # #     columns = [
# # #         "product_id",
# # #         "total_purchased",
# # #         "remaining_inventory",
# # #         "total_purchase_cost",
# # #         "total_sold",
# # #         "total_sale_revenue",
# # #         "profit_all_time",
# # #         "avg_flip_amount",
# # #         "min_purchase_price",
# # #         "best_purchase_date",
# # #         "best_purchase_contact",
# # #         "best_customer_id",
# # #         "best_customer_total_purchased",
# # #         "best_sale_amount",
# # #         "best_sale_date"
# # #     ]
# # #
# # #     df = pd.DataFrame(results, columns=columns)
# # #     df.to_excel(REPORT_FILE, sheet_name=sheet_name, index=False)
# # #
# # #
# # #
# # # dataaddagain.generate_excel_report(REPORT_FILE, False)
# # # calculate_and_update_inventory_stats('inventory analytics')
# # # create_pivot_with_xlwings(REPORT_FILE, pivot_sheet_name="Sheet1", raw_sheet_name="Pivot1")
# #
# # import os
# # import sys
# # import sqlite3
# # import pandas as pd
# # import xlwings as xw
# # import datetime
# #
# # print("Pandas imported successfully!")
# #
# # # ===================== CONFIGURATION =====================
# # if getattr(sys, 'frozen', False):
# #     base_dir = os.path.dirname(sys.executable)
# # else:
# #     base_dir = os.path.dirname(__file__)
# # print("Base directory:", base_dir)
# #
# # DB_FILE = os.path.join(base_dir, "business_tracker.db")
# # REPORT_FILE = os.path.join(base_dir, "full_analytics_report.xlsx")
# #
# #
# # # ===================== DATABASE HELPER FUNCTION =====================
# # def execute_query(query, params=(), fetch=False):
# #     conn = sqlite3.connect(DB_FILE)
# #     cursor = conn.cursor()
# #     cursor.execute(query, params)
# #     data = cursor.fetchall() if fetch else None
# #     conn.commit()
# #     conn.close()
# #     return data
# #
# #
# # # ===================== External Functions and Globals =====================
# # # We assume that dataaddagain and UpdateDatabaseWithCalculatedFields are modules you maintain.
# # import dataaddagain
# # import UpdateDatabaseWithCalculatedFields
# #
# #
# # # ===================== Function: Calculate & Export Inventory Stats =====================
# # def calculate_and_update_inventory_stats(sheet_name):
# #     """
# #     Calculates inventory statistics for every product using precomputed data from
# #     Offer_Product and CustomerSale_Product, stores the results in a DataFrame,
# #     and exports that DataFrame to an Excel file (deleting any existing file first).
# #     The output is written to a sheet with the name provided in sheet_name.
# #     The query returns 15 columns: product_id, total_purchased, remaining_inventory,
# #     total_purchase_cost, total_sold, total_sale_revenue, profit_all_time, avg_flip_amount,
# #     min_purchase_price, best_purchase_date, best_purchase_contact, best_customer_id,
# #     best_customer_total_purchased, best_sale_amount, best_sale_date.
# #     """
# #     # Get all products (keys as strings)
# #     prod_dict = UpdateDatabaseWithCalculatedFields.get_products()
# #     results = []
# #
# #     query_inventory_statistics = """
# #         SELECT
# #             ? AS product_id,
# #             (SELECT COALESCE(SUM(quantity_purchased), 0)
# #              FROM Offer_Product WHERE product_id = ?),
# #             ((SELECT COALESCE(SUM(quantity_purchased), 0)
# #               FROM Offer_Product WHERE product_id = ?)
# #              - (SELECT COALESCE(SUM(quantity_sold), 0)
# #                 FROM CustomerSale_Product WHERE product_id = ?)),
# #             (SELECT COALESCE(SUM(quantity_purchased * purchase_price_per_unit), 0)
# #              FROM Offer_Product WHERE product_id = ?),
# #             (SELECT COALESCE(SUM(quantity_sold), 0)
# #              FROM CustomerSale_Product WHERE product_id = ?),
# #             (SELECT COALESCE(SUM(quantity_sold * sell_price_per_unit), 0)
# #              FROM CustomerSale_Product WHERE product_id = ?),
# #             ((SELECT COALESCE(SUM(quantity_sold * sell_price_per_unit), 0)
# #               FROM CustomerSale_Product WHERE product_id = ?)
# #              - (SELECT COALESCE(SUM(quantity_purchased * purchase_price_per_unit), 0)
# #                 FROM Offer_Product WHERE product_id = ?)),
# #             (SELECT AVG(sell_price_per_unit)
# #              FROM CustomerSale_Product WHERE product_id = ?),
# #             (SELECT MIN(purchase_price_per_unit)
# #              FROM Offer_Product WHERE product_id = ?),
# #             (SELECT o.offer_end_date
# #              FROM Offer o JOIN Offer_Product op ON o.id = op.offer_id
# #              WHERE op.product_id = ? ORDER BY op.purchase_price_per_unit ASC LIMIT 1),
# #             (SELECT c.name
# #              FROM Contact c JOIN Offer o ON o.contact_id = c.id JOIN Offer_Product op ON o.id = op.offer_id
# #              WHERE op.product_id = ? ORDER BY op.purchase_price_per_unit ASC LIMIT 1),
# #             (SELECT customer_id
# #              FROM (
# #                  SELECT cs.contact_id AS customer_id, SUM(csp.quantity_sold) AS total_qty
# #                  FROM CustomerSale_Product csp
# #                  JOIN CustomerSale cs ON cs.id = csp.customer_sale_id
# #                  WHERE csp.product_id = ?
# #                  GROUP BY cs.contact_id
# #                  ORDER BY total_qty DESC LIMIT 1
# #              )),
# #             (SELECT total_qty
# #              FROM (
# #                  SELECT cs.contact_id AS customer_id, SUM(csp.quantity_sold) AS total_qty
# #                  FROM CustomerSale_Product csp
# #                  JOIN CustomerSale cs ON cs.id = csp.customer_sale_id
# #                  WHERE csp.product_id = ?
# #                  GROUP BY cs.contact_id
# #                  ORDER BY total_qty DESC LIMIT 1
# #              )),
# #             (SELECT MAX(sell_price_per_unit)
# #              FROM CustomerSale_Product WHERE product_id = ?),
# #             (SELECT cs.sell_date
# #              FROM CustomerSale cs JOIN CustomerSale_Product csp ON cs.id = csp.customer_sale_id
# #              WHERE csp.product_id = ? ORDER BY csp.sell_price_per_unit DESC LIMIT 1)
# #     """
# #     # Each product_id is supplied 17 times
# #     for prod_id in prod_dict.keys():
# #         params = tuple([prod_id] * 17)
# #         row = execute_query(query_inventory_statistics, params, fetch=True)
# #         if row:
# #             results.append(row[0])
# #
# #     columns = [
# #         "product_id",
# #         "total_purchased",
# #         "remaining_inventory",
# #         "total_purchase_cost",
# #         "total_sold",
# #         "total_sale_revenue",
# #         "profit_all_time",
# #         "avg_flip_amount",
# #         "min_purchase_price",
# #         "best_purchase_date",
# #         "best_purchase_contact",
# #         "best_customer_id",
# #         "best_customer_total_purchased",
# #         "best_sale_amount",
# #         "best_sale_date"
# #     ]
# #     df = pd.DataFrame(results, columns=columns)
# #     # Write the DataFrame to the specified sheet in the Excel file.
# #     # This will overwrite any existing file.
# #     df.to_excel(REPORT_FILE, sheet_name=sheet_name, index=False)
# #     print(f"Inventory statistics written to sheet '{sheet_name}' in '{REPORT_FILE}'.")
# #
# #
# # # ===================== Function: Create Pivot Table Using xlwings =====================
# # def create_pivot_with_xlwings(excel_file, raw_sheet_name, pivot_sheet_name):
# #     """
# #     Opens the Excel workbook (with raw data already on raw_sheet_name) and creates a pivot table
# #     on a new sheet named pivot_sheet_name.
# #     If raw_sheet_name is not found, falls back to "Sheet1".
# #     """
# #     app = xw.App(visible=False)
# #     wb = app.books.open(excel_file)
# #     available_sheets = [s.name for s in wb.sheets]
# #     print("Available sheets:", available_sheets)
# #     if raw_sheet_name in available_sheets:
# #         raw_sheet = wb.sheets[raw_sheet_name]
# #     elif "Sheet1" in available_sheets:
# #         print(f"Sheet '{raw_sheet_name}' not found. Falling back to 'Sheet1'.")
# #         raw_sheet = wb.sheets["Sheet1"]
# #     else:
# #         print(f"Error: Neither '{raw_sheet_name}' nor 'Sheet1' were found!")
# #         wb.close()
# #         app.quit()
# #         return
# #
# #     if pivot_sheet_name in available_sheets:
# #         pivot_sheet = wb.sheets[pivot_sheet_name]
# #         pivot_sheet.clear()
# #     else:
# #         pivot_sheet = wb.sheets.add(pivot_sheet_name)
# #
# #     source_range = raw_sheet.range("A1").expand()
# #     pivot_table = pivot_sheet.api.PivotTableWizard(
# #         SourceType=1,  # xlDatabase
# #         SourceData=source_range.api,
# #         TableDestination=pivot_sheet.range("A3").api,
# #         TableName="PivotTable1"
# #     )
# #
# #     # Configure pivot table fields.
# #     # Example: Group by CustomerName, Sell_Date, and Sold_ProductName; sum TotalSalePrice.
# #     pivot_table.PivotFields("CustomerName").Orientation = 1  # Row field
# #     pivot_table.PivotFields("Sell_Date").Orientation = 1  # Row field
# #     pivot_table.PivotFields("Sold_ProductName").Orientation = 1  # Row field
# #
# #     data_field = pivot_table.PivotFields("TotalSalePrice")
# #     data_field.Orientation = 4  # Data field
# #     data_field.Function = -4157  # xlSum
# #     data_field.Name = "Sum of TotalSalePrice"
# #
# #     wb.save()
# #     wb.close()
# #     app.quit()
# #     print(f"Pivot table created on sheet '{pivot_sheet_name}' in '{excel_file}'.")
# #
# #
# # # ===================== Function: Get Offer and Sale ID Lists =====================
# # # def get_offer_and_sale_ids(session_only=True):
# # #     """
# # #     Returns a tuple (offer_id_list, sale_id_list) as comma-separated strings.
# # #     If session_only is True, uses UpdateDatabaseWithCalculatedFields.session_offer_ids
# # #     and session_sale_ids. Otherwise, queries the database for all IDs.
# # #     """
# # #     if session_only:
# # #         offer_id_list = ",".join(map(str, UpdateDatabaseWithCalculatedFields.session_offer_ids))
# # #         sale_id_list = ",".join(map(str,
# # #                                     UpdateDatabaseWithCalculatedFields.session_sale_ids)) if UpdateDatabaseWithCalculatedFields.session_sale_ids else "0"
# # #     else:
# # #         all_offers = execute_query("SELECT id FROM Offer", fetch=True)
# # #         offer_ids = [str(row[0]) for row in all_offers]
# # #         offer_id_list = ",".join(offer_ids)
# # #         all_sales = execute_query("SELECT id FROM CustomerSale", fetch=True)
# # #         sale_ids = [str(row[0]) for row in all_sales]
# # #         sale_id_list = ",".join(sale_ids) if sale_ids else "0"
# # #     return offer_id_list, sale_id_list
# #
# #
# # # ===================== MAIN EXECUTION =====================
# # if __name__ == "__main__":
# #     # 1) Generate raw data report.
# #     # This call is assumed to create the raw data sheet in REPORT_FILE (e.g., "Sheet1")
# #     dataaddagain.generate_excel_report(REPORT_FILE, False)
# #
# #     # 2) Calculate and update inventory statistics; write results to sheet "Inventory Analytics"
# #     calculate_and_update_inventory_stats("Inventory Analytics")
# #
# #     # 3) Get offer and sale ID lists (session-only format)
# #     # offer_id_list, sale_id_list = get_offer_and_sale_ids(session_only=True)
# #     # print("Offer ID List:", offer_id_list)
# #     # print("Sale ID List:", sale_id_list)
# #
# #     # 4) Create a pivot table from the raw data.
# #     # Here we assume that dataaddagain.generate_excel_report has created the raw data in "Sheet1"
# #     create_pivot_with_xlwings(REPORT_FILE, raw_sheet_name="Sheet1", pivot_sheet_name="Pivot1")
# #
# #     print("Process complete! Excel file with pivot table created at:", REPORT_FILE)

import os
import sys
import sqlite3
import pandas as pd
import xlwings as xw
import datetime
import UpdateDatabaseWithCalculatedFields
import DataInsertGUI

print("Pandas imported successfully!")

# ===================== CONFIGURATION =====================
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(__file__)
print("Base directory:", base_dir)

DB_FILE = os.path.join(base_dir, "business_tracker.db")
REPORT_FILE = os.path.join(base_dir, "full_analytics_report.xlsx")


def execute_query(query, params=(), fetch=False):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(query, params)
    data = cursor.fetchall() if fetch else None
    conn.commit()
    conn.close()
    return data


# ===================== calculate_and_update_inventory_stats =====================
def calculate_and_update_inventory_stats(sheet_name):
    """
    Calculates inventory statistics for every product using precomputed data from
    Offer_Product and CustomerSale_Product, and writes the results to an Excel sheet.
    Instead of deleting the entire file, it appends a new sheet named `sheet_name`
    to the existing workbook (if it exists).
    """
    prod_dict = UpdateDatabaseWithCalculatedFields.get_products()
    results = []
    query_inventory_statistics = """
        SELECT 
            ? AS product_id,
            (SELECT COALESCE(SUM(quantity_purchased), 0)
             FROM Offer_Product WHERE product_id = ?),
            ((SELECT COALESCE(SUM(quantity_purchased), 0)
              FROM Offer_Product WHERE product_id = ?)
             - (SELECT COALESCE(SUM(quantity_sold), 0)
                FROM CustomerSale_Product WHERE product_id = ?)),
            (SELECT COALESCE(SUM(quantity_purchased * purchase_price_per_unit), 0)
             FROM Offer_Product WHERE product_id = ?),
            (SELECT COALESCE(SUM(quantity_sold), 0)
             FROM CustomerSale_Product WHERE product_id = ?),
            (SELECT COALESCE(SUM(quantity_sold * sell_price_per_unit), 0)
             FROM CustomerSale_Product WHERE product_id = ?),
            ((SELECT COALESCE(SUM(quantity_sold * sell_price_per_unit), 0)
              FROM CustomerSale_Product WHERE product_id = ?)
             - (SELECT COALESCE(SUM(quantity_purchased * purchase_price_per_unit), 0)
                FROM Offer_Product WHERE product_id = ?)),
            (SELECT AVG(sell_price_per_unit)
             FROM CustomerSale_Product WHERE product_id = ?),
            (SELECT MIN(purchase_price_per_unit)
             FROM Offer_Product WHERE product_id = ?),
            (SELECT o.offer_end_date
             FROM Offer o JOIN Offer_Product op ON o.id = op.offer_id
             WHERE op.product_id = ? ORDER BY op.purchase_price_per_unit ASC LIMIT 1),
            (SELECT c.name
             FROM Contact c JOIN Offer o ON o.contact_id = c.id JOIN Offer_Product op ON o.id = op.offer_id
             WHERE op.product_id = ? ORDER BY op.purchase_price_per_unit ASC LIMIT 1),
            (SELECT customer_id
             FROM (
                 SELECT cs.contact_id AS customer_id, SUM(csp.quantity_sold) AS total_qty
                 FROM CustomerSale_Product csp
                 JOIN CustomerSale cs ON cs.id = csp.customer_sale_id
                 WHERE csp.product_id = ?
                 GROUP BY cs.contact_id
                 ORDER BY total_qty DESC LIMIT 1
             )),
            (SELECT total_qty
             FROM (
                 SELECT cs.contact_id AS customer_id, SUM(csp.quantity_sold) AS total_qty
                 FROM CustomerSale_Product csp
                 JOIN CustomerSale cs ON cs.id = csp.customer_sale_id
                 WHERE csp.product_id = ?
                 GROUP BY cs.contact_id
                 ORDER BY total_qty DESC LIMIT 1
             )),
            (SELECT MAX(sell_price_per_unit)
             FROM CustomerSale_Product WHERE product_id = ?),
            (SELECT cs.sell_date
             FROM CustomerSale cs JOIN CustomerSale_Product csp ON cs.id = csp.customer_sale_id
             WHERE csp.product_id = ? ORDER BY csp.sell_price_per_unit DESC LIMIT 1)
    """
    for prod_id in prod_dict.keys():
        params = tuple([prod_id] * 17)
        row = execute_query(query_inventory_statistics, params, fetch=True)
        if row:
            results.append(row[0])
    columns = [
        "product_id",
        "total_purchased",
        "remaining_inventory",
        "total_purchase_cost",
        "total_sold",
        "total_sale_revenue",
        "profit_all_time",
        "avg_flip_amount",
        "min_purchase_price",
        "best_purchase_date",
        "best_purchase_contact",
        "best_customer_id",
        "best_customer_total_purchased",
        "best_sale_amount",
        "best_sale_date"
    ]
    df = pd.DataFrame(results, columns=columns)

    # Append to the existing workbook if it exists; otherwise, create new.
    if os.path.exists(REPORT_FILE):
        with pd.ExcelWriter(REPORT_FILE, engine="openpyxl", mode="a") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        df.to_excel(REPORT_FILE, sheet_name=sheet_name, index=False)

    print(f"Inventory statistics written to sheet '{sheet_name}' in '{REPORT_FILE}'.")


# ===================== create_pivot_with_xlwings =====================
def create_second_pivot_with_xlwings(excel_file, raw_sheet_name, pivot_sheet_name):
    """
    Opens the Excel workbook (with raw data already on the sheet named by raw_sheet_name)
    and creates a second pivot table on a new sheet (pivot_sheet_name). The pivot table is built
    dynamically by extracting header names from the raw data (assumed to be in row 1) and then:

      - Using "Offer_ID", "ProductName", and "Sale_Unit" as row fields.
      - Summing "Quantity_Sold" and "TotalSalePrice" as data fields.

    If raw_sheet_name is not found, it falls back to "Sheet1".
    """
    import xlwings as xw

    # Start Excel in visible mode for debugging (set to False when finished)
    app = xw.App(visible=True, add_book=False)
    wb = app.books.open(excel_file)

    available_sheets = [s.name for s in wb.sheets]
    print("Available sheets:", available_sheets)

    # Get the raw data sheet
    if raw_sheet_name in available_sheets:
        raw_sheet = wb.sheets[raw_sheet_name]
    elif "Sheet1" in available_sheets:
        print(f"Sheet '{raw_sheet_name}' not found. Falling back to 'Sheet1'.")
        raw_sheet = wb.sheets["Sheet1"]
    else:
        print(f"Error: Neither '{raw_sheet_name}' nor 'Sheet1' were found!")
        wb.close()
        app.quit()
        return

    # Extract headers from row 1 (A1:V1) and build a dictionary mapping lowercase header names to exact values.
    header_cells = raw_sheet.range("A1:V1")
    headers = [str(cell.value).strip() for cell in header_cells if cell.value is not None]
    print("Extracted headers:", headers)
    header_map = {h.lower(): h for h in headers}

    # Define the desired fields for pivot table 2 (using lowercase keys).
    desired_row_fields = ["offer_id", "productname", "sale_unit"]
    desired_data_fields = ["quantity_sold", "totalsaleprice"]

    # Verify that the desired fields exist.
    for field in desired_row_fields + desired_data_fields:
        if field not in header_map:
            print(f"Error: Required field '{field}' not found in headers!")
            wb.close()
            app.quit()
            return

    # Create or clear the pivot sheet.
    if pivot_sheet_name in available_sheets:
        pivot_sheet = wb.sheets[pivot_sheet_name]
        pivot_sheet.clear()
    else:
        pivot_sheet = wb.sheets.add(pivot_sheet_name)

    # Determine the source range dynamically.
    # Using expand("table")—if this returns too few rows, consider using expand() or a manual range.
    source_range = raw_sheet.range("A1").expand("table")
    print("Source range address:", source_range.address)
    # Build an A1-style reference string.
    range_string = f"'{raw_sheet.name}'!$A$1:$V$4"
    print("Source range string:", range_string)

    source_range = raw_sheet.range("A1").expand()

    pivot_table = pivot_sheet.api.PivotTableWizard(
        SourceType=1,  # xlDatabase
        SourceData=source_range.api,
        TableDestination=pivot_sheet.range("A3").api,
        TableName="PivotTable2"  # Use a unique table name.
    )

    # Configure row fields dynamically.
    for field in desired_row_fields:
        actual_field = header_map[field]
        pivot_table.PivotFields(actual_field).Orientation = 1  # Row field

    # Add the first data field: Sum of Quantity_Sold
    qty_field = pivot_table.PivotFields(header_map["quantity_sold"])
    qty_field.Orientation = 4  # Data field
    qty_field.Function = -4157  # xlSum
    qty_field.Name = "Sum of Quantity_Sold"

    # Add the second data field: Sum of TotalSalePrice
    price_field = pivot_table.PivotFields(header_map["totalsaleprice"])
    price_field.Orientation = 4  # Data field
    price_field.Function = -4157  # xlSum
    price_field.Name = "Sum of TotalSalePrice"

    wb.save()
    wb.close()
    app.quit()
    print(f"Second pivot table created on sheet '{pivot_sheet_name}' in '{excel_file}'.")


def create_pivot_with_xlwings(excel_file, raw_sheet_name, pivot_sheet_name):
    """
    Opens the Excel workbook and creates a pivot table on a new sheet named pivot_sheet_name.
    It uses data from the sheet specified by raw_sheet_name (or falls back to "Sheet1").
    The field names for the pivot table are extracted dynamically from the header row.
    """
    import xlwings as xw

    app = xw.App(visible=True, add_book=False)
    wb = app.books.open(excel_file)

    available_sheets = [s.name for s in wb.sheets]
    print("Available sheets:", available_sheets)

    if raw_sheet_name in available_sheets:
        raw_sheet = wb.sheets[raw_sheet_name]
    elif "Sheet1" in available_sheets:
        print(f"Sheet '{raw_sheet_name}' not found. Falling back to 'Sheet1'.")
        raw_sheet = wb.sheets["Sheet1"]
    else:
        print(f"Error: Neither '{raw_sheet_name}' nor 'Sheet1' were found!")
        wb.close()
        app.quit()
        return

    # Extract headers from the first row (A1:V1)
    header_cells = raw_sheet.range("A1:V1")
    headers = [str(cell.value).strip() for cell in header_cells if cell.value is not None]
    print("Extracted headers:", headers)
    header_cells = raw_sheet.range("A1:V1")
    headers = [str(cell.value).strip() for cell in header_cells if cell.value is not None]
    print("Raw headers (pre-clean):", headers)
    clean_headers = [h.strip("'") for h in headers]
    print("Clean headers:", clean_headers)
    # Build a dictionary mapping lowercase header to actual header text
    header_map = {h.lower(): h for h in headers}

    # For example, desired pivot fields:
    desired_row_fields = ["customername", "sell_date", "sold_productname"]
    desired_data_field = "totalsaleprice"

    # Verify that the desired fields exist:
    for field in desired_row_fields + [desired_data_field]:
        if field not in header_map:
            print(f"Error: Required field '{field}' not found in headers!")
            wb.close()
            app.quit()
            return

    # Create or clear the pivot sheet
    if pivot_sheet_name in available_sheets:
        pivot_sheet = wb.sheets[pivot_sheet_name]
        pivot_sheet.clear()
    else:
        pivot_sheet = wb.sheets.add(pivot_sheet_name)

    # Determine the source range using a dynamic expansion
    source_range = raw_sheet.range("A1").expand("table")
    range_string = f"'{raw_sheet.name}'!$A$1:$V$4"

    print("Source range address:", range_string)
    source_range = raw_sheet.range("A1").expand()

    pivot_table = pivot_sheet.api.PivotTableWizard(
        SourceType=1,  # xlDatabase
        SourceData=source_range.api,
        TableDestination=pivot_sheet.range("A3").api,
        TableName="PivotTable1"  # Use a unique table name.
    )

    # Configure row fields using the dynamic header lookup:
    for field in desired_row_fields:
        actual_field = header_map[field]
        pivot_table.PivotFields(actual_field).Orientation = 1  # Row field

    # Configure data field:
    actual_data_field = header_map[desired_data_field]
    data_field = pivot_table.PivotFields(actual_data_field)
    data_field.Orientation = 4  # Data field
    data_field.Function = -4157  # xlSum
    data_field.Name = f"Sum of {actual_data_field}"

    wb.save()
    wb.close()
    app.quit()
    print(f"Pivot table created on sheet '{pivot_sheet_name}' in '{excel_file}'.")


# ===================== MAIN EXECUTION =====================
if __name__ == "__main__":
    # 1) Generate raw data report; this should create a sheet named "Sheet1" if dataaddagain does so.
    dataaddagain.generate_excel_report(REPORT_FILE, False)

    # 2) Append inventory statistics to a new sheet named "Inventory Analytics"
    calculate_and_update_inventory_stats("Inventory Analytics")

    # 3) Get offer and sale ID lists (session-only)

    # 4) Create a pivot table using raw data from "Sheet1".
    create_pivot_with_xlwings(REPORT_FILE, raw_sheet_name="Sheet1", pivot_sheet_name="Pivot1")
    create_second_pivot_with_xlwings(REPORT_FILE, raw_sheet_name="Sheet1", pivot_sheet_name="Pivot2")
    print("Process complete! Excel file with pivot table created at:", REPORT_FILE)
