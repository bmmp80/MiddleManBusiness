# import os
# import sqlite3
# import tkinter as tk
# from tkinter import ttk, messagebox, Toplevel
# import pandas as pd
# import numpy as np
# import datetime
# import sys
# import UpdateDatabaseWithCalculatedFields
#
# print("Using Python:", sys.executable)
# print("Python version:", sys.version)
# print("sys.path:", sys.path)
#
# import pandas as pd
# print("Pandas imported successfully!")
# import time
# import random
# import re
#
# # ===================== CONFIGURATION =====================
# if getattr(sys, 'frozen', False):
#     base_dir = os.path.dirname(sys.executable)
# else:
#     base_dir = os.path.dirname(__file__)
# print(base_dir)
# DB_FILE = os.path.join(base_dir, "business_tracker.db")
# REPORT_FILE = os.path.join(base_dir, "financial_session_report.xlsx")
#
# # ===================== GLOBAL VARIABLES =====================
# session_offer_ids = []  # New offer IDs created this session
# session_sale_ids = []   # New sale IDs created this session
#
# suppliers = {}
# customers = {}
# products = {}
# offers = {}
#
# # For the Add Offer Tab product widgets:
# qty_entries = {}              # product_id -> Entry widget for quantity
# product_vars = {}             # product_id -> BooleanVar for product selection
# purchase_price_entries = {}   # product_id -> Entry widget for purchase price
# offer_unit_entries = {}       # product_id -> Entry widget for unit
#
# # For the Customer Sale Tab:
# sale_product_vars = {}        # product_id -> BooleanVar for sale product selection
# sale_qty_entries = {}         # product_id -> Entry widget for sale quantity
# sale_price_entries = {}       # product_id -> Entry widget for sale price
# sale_unit_entries = {}        # product_id -> Entry widget for sale unit
#
# # For the Edit Offer Tab:
# edit_qty_entries = {}              # product_id -> Entry widget for quantity (edit tab)
# edit_product_vars = {}             # product_id -> BooleanVar (edit tab)
# edit_purchase_price_entries = {}   # product_id -> Entry widget for purchase price (edit tab)
# edit_offer_unit_entries = {}       # product_id -> Entry widget for unit (edit tab)
# # Global dictionary to preserve manual changes when refreshing the edit product list.
# edit_offer_state = {}  # product_id -> (checked, qty, price, unit)
#
# DEBUG = True  # Enable debug prints
#
# # ===================== DATABASE HELPER FUNCTIONS =====================
# def execute_query(query, params=(), fetch=False):
#     conn = sqlite3.connect(DB_FILE)
#     cursor = conn.cursor()
#     cursor.execute(query, params)
#     data = cursor.fetchall() if fetch else None
#     conn.commit()
#     conn.close()
#     return data
#
# def get_suppliers():
#     return {str(c[0]): c[1] for c in execute_query(
#         "SELECT id, name FROM Contact WHERE type='supplier'", fetch=True)}
#
# def get_customers():
#     return {str(c[0]): c[1] for c in execute_query(
#         "SELECT id, name FROM Contact WHERE type='customer'", fetch=True)}
#
# def get_products():
#     return {str(p[0]): p[1] for p in execute_query(
#         "SELECT id, name FROM Product", fetch=True)}
#
# def get_offers():
#     rows = execute_query(
#         """
#         SELECT
#             Offer.id,
#             Contact.name,
#             (
#                 SELECT COALESCE(SUM(op.quantity_purchased * op.purchase_price_per_unit), 0)
#                 FROM Offer_Product op
#                 WHERE op.offer_id = Offer.id
#             ) AS Investment
#         FROM Offer
#         JOIN Contact ON Offer.contact_id = Contact.id
#         """, fetch=True)
#     # Build strings like "2 - SoCal ($1776.0)"; keys kept as numbers.
#     return {o[0]: f"{o[0]} - {o[1]} (${o[2]})" for o in rows}
#
# def get_products_by_offer(offer_id):
#     results = execute_query(
#         """
#         SELECT p.id, p.name, op.quantity_purchased
#         FROM Offer_Product op
#         JOIN Product p ON op.product_id = p.id
#         WHERE op.offer_id = ?
#         """, (int(offer_id),), fetch=True)
#     return {r[0]: (r[1], r[2]) for r in results}
#
# # ===================== UNIT CONVERSION FUNCTION =====================
# def convert_units(value, from_unit, to_unit):
#     """Converts value from from_unit to to_unit using defined conversion rates."""
#     from_unit = from_unit.lower().strip()
#     to_unit = to_unit.lower().strip()
#     conversion_rates = {
#         ("oz", "g"): 28,
#         ("g", "oz"): 1 / 28,
#         ("lb", "kg"): 0.453592,
#         ("kg", "lb"): 2.20462,
#         ("g", "kg"): 0.001,
#         ("kg", "g"): 1000,
#     }
#     if from_unit == to_unit:
#         return value
#     if (from_unit, to_unit) in conversion_rates:
#         return value * conversion_rates[(from_unit, to_unit)]
#     else:
#         print(f"Warning: No conversion rate from {from_unit} to {to_unit}. Using original value.")
#         return value
#
# # ===================== REFRESH FUNCTIONS =====================
# def refresh_dropdowns():
#     global suppliers, customers, products, offers
#     suppliers = get_suppliers()
#     customers = get_customers()
#     products = get_products()
#     offers = get_offers()
#     if DEBUG:
#         print('refreshed dropdowns')
#         print('offers:', offers)
#     contact_dropdown["values"] = list(suppliers.values())
#     customer_dropdown["values"] = list(customers.values())
#     offer_dropdown["values"] = list(offers.values())
#     try:
#         edit_offer_dropdown["values"] = list(offers.values())
#         edit_contact_dropdown["values"] = list(suppliers.values())
#     except Exception:
#         pass
#     refresh_product_list()
#     refresh_sale_product_list()
#     # If an offer is already selected in Edit Offer tab, refresh its details.
#     selected = edit_offer_dropdown.get().strip()
#     # CHANGE - might have to have offer drop down same logic as above
#     if selected:
#         load_offer_details(None)
#
#
# def refresh_product_list():
#     global products, qty_entries, product_vars, purchase_price_entries, offer_unit_entries
#     products = get_products()
#     for widget in product_frame.winfo_children():
#         widget.destroy()
#     qty_entries = {}
#     product_vars = {}
#     purchase_price_entries = {}
#     offer_unit_entries = {}
#     tk.Label(product_frame, text="Products:").grid(row=0, column=0, sticky="w")
#     tk.Label(product_frame, text="Qty").grid(row=0, column=1)
#     tk.Label(product_frame, text="Unit Price").grid(row=0, column=2)
#     tk.Label(product_frame, text="Unit").grid(row=0, column=3)
#     row_index = 1
#     for product_id, product_name in products.items():
#         var = tk.BooleanVar()
#         chk = tk.Checkbutton(product_frame, text=product_name, variable=var,
#                              command=lambda pid=product_id, v=var: toggle_quantity_field(pid, v))
#         chk.grid(row=row_index, column=0, sticky="w", padx=5, pady=2)
#         qty_entry = tk.Entry(product_frame, state="disabled", width=5)
#         qty_entry.grid(row=row_index, column=1, padx=5)
#         price_entry = tk.Entry(product_frame, state="disabled", width=7)
#         price_entry.grid(row=row_index, column=2, padx=5)
#         unit_entry = tk.Entry(product_frame, state="disabled", width=5)
#         unit_entry.grid(row=row_index, column=3, padx=5)
#         product_vars[product_id] = var
#         qty_entries[product_id] = qty_entry
#         purchase_price_entries[product_id] = price_entry
#         offer_unit_entries[product_id] = unit_entry
#         row_index += 1
#
# def refresh_sale_product_list():
#     # For the Customer Sale tab: list only products that are associated with the selected offer.
#
#     try:
#         # For the Customer Sale tab: list only products that are associated with the selected offer.
#         selected_offer_str = offer_dropdown.get()
#         offer_id = int(selected_offer_str.split()[0])
#     except (ValueError, IndexError) as e:
#         if DEBUG:
#             print("refresh_sale_product_list: No valid offer selected:", e)
#         return
#
#     # offer_id = next((k for k, v in offers.items() if v == selected_offer), None)
#     # Clear the existing frame
#     for widget in sale_product_frame.winfo_children():
#         widget.destroy()
#     sale_product_vars.clear()
#     sale_qty_entries.clear()
#     sale_price_entries.clear()
#     sale_unit_entries.clear()
#
#     if not offer_id:
#         return  # No valid offer selected
#
#     # Use get_products_by_offer to list only products linked to this offer
#     offer_products = get_products_by_offer(offer_id)
#     # offer_products is a dict: { product_id: (product_name, purchased_qty), ... }
#
#     # Define buddy units for display
#     buddy_map = {
#         "oz": "g",
#         "g": "oz",
#         "lb": "kg",
#         "kg": "lb"
#     }
#
#     tk.Label(sale_product_frame, text="Products for Sale:").grid(row=0, column=0, sticky="w")
#     tk.Label(sale_product_frame, text="Qty").grid(row=0, column=1)
#     tk.Label(sale_product_frame, text="Unit Sell Price").grid(row=0, column=2)
#     tk.Label(sale_product_frame, text="Unit").grid(row=0, column=3)
#
#     row_index = 1
#     for product_id, (product_name, _) in offer_products.items():
#         # 1) Find the product's purchase quantity & unit from Offer_Product
#         purchase_data = execute_query(
#             "SELECT quantity_purchased, unit FROM Offer_Product WHERE offer_id=? AND product_id=?",
#             (offer_id, product_id),
#             fetch=True
#         )
#         if not purchase_data:
#             # Shouldn't happen if get_products_by_offer found it, but just in case:
#             continue
#         purchased_qty, purchase_unit = purchase_data[0]
#
#         # 2) Sum up how much has been sold for this product so far (converting units to purchase_unit)
#         sales_data = execute_query(
#             """
#             SELECT csp.quantity_sold, u.unit_name
#             FROM CustomerSale_Product csp
#             JOIN CustomerSale cs ON cs.id = csp.customer_sale_id
#             JOIN Unit u ON u.unit_id = csp.unit_id
#             WHERE cs.offer_id=? AND csp.product_id=?
#             """,
#             (offer_id, product_id),
#             fetch=True
#         )
#         total_sold_converted = 0.0
#         for sold_qty, sold_unit in sales_data:
#             total_sold_converted += convert_units(sold_qty, sold_unit, purchase_unit)
#
#         # 3) Compute remaining in the purchase unit
#         remaining = purchased_qty - total_sold_converted
#         if remaining < 0:
#             remaining = 0.0  # Just in case
#
#         # 4) Convert to a "buddy" unit if applicable (e.g. oz ↔ g, lb ↔ kg)
#         purchase_unit_lower = purchase_unit.lower().strip()
#         buddy_unit = buddy_map.get(purchase_unit_lower, None)
#         if buddy_unit:
#             # Convert the remaining from purchase_unit to buddy_unit
#             buddy_qty = convert_units(remaining, purchase_unit_lower, buddy_unit)
#             # e.g. "dd (remaining: 1.00oz / 28.35g)"
#             display_text = f"{product_name} (remaining: {remaining:.2f}{purchase_unit} / {buddy_qty:.2f}{buddy_unit})"
#         else:
#             # If there's no buddy unit, just show the remaining in the purchase unit
#             display_text = f"{product_name} (remaining: {remaining:.2f}{purchase_unit})"
#
#         # 5) Build the row with a checkbox and the relevant Entry fields
#         var = tk.BooleanVar()
#         chk = tk.Checkbutton(sale_product_frame, text=display_text, variable=var)
#         chk.grid(row=row_index, column=0, sticky="w", padx=5, pady=2)
#
#         qty_entry = tk.Entry(sale_product_frame, state="disabled", width=5)
#         qty_entry.grid(row=row_index, column=1, padx=5)
#
#         price_entry = tk.Entry(sale_product_frame, state="disabled", width=7)
#         price_entry.grid(row=row_index, column=2, padx=5)
#
#         unit_entry = tk.Entry(sale_product_frame, state="disabled", width=5)
#         unit_entry.grid(row=row_index, column=3, padx=5)
#
#         # Enable/disable the entries based on checkbox toggle
#         def on_toggle(pid=product_id, v=var, q_entry=qty_entry, p_entry=price_entry, u_entry=unit_entry):
#             if v.get():
#                 q_entry.config(state="normal")
#                 p_entry.config(state="normal")
#                 u_entry.config(state="normal")
#             else:
#                 q_entry.config(state="disabled")
#                 p_entry.config(state="disabled")
#                 u_entry.config(state="disabled")
#
#         chk.config(command=on_toggle)
#
#         # Store references so we can read them later on submission
#         sale_product_vars[product_id] = var
#         sale_qty_entries[product_id] = qty_entry
#         sale_price_entries[product_id] = price_entry
#         sale_unit_entries[product_id] = unit_entry
#
#         row_index += 1
#
#
# def build_sale_product_list():
#     refresh_sale_product_list()
#
# def refresh_customer_sale_products(event=None):
#     refresh_sale_product_list()
#
# # ===================== TOGGLE QUANTITY FUNCTIONS =====================
# def toggle_quantity_field(product_id, var):
#     if var.get():
#         qty_entries[product_id].config(state="normal")
#         purchase_price_entries[product_id].config(state="normal")
#         offer_unit_entries[product_id].config(state="normal")
#     else:
#         qty_entries[product_id].config(state="disabled")
#         purchase_price_entries[product_id].config(state="disabled")
#         offer_unit_entries[product_id].config(state="disabled")
#
# def toggle_edit_quantity_field(product_id, var):
#     if var.get():
#         edit_qty_entries[product_id].config(state="normal")
#         edit_purchase_price_entries[product_id].config(state="normal")
#         edit_offer_unit_entries[product_id].config(state="normal")
#     else:
#         edit_qty_entries[product_id].config(state="disabled")
#         edit_purchase_price_entries[product_id].config(state="disabled")
#         edit_offer_unit_entries[product_id].config(state="disabled")
#
# # ===================== ADD/UPDATE FUNCTIONS =====================
#
# #change to append offer to session offer IDs
# def add_offer():
#     global suppliers
#     suppliers = get_suppliers()
#     contact_name = contact_dropdown.get()
#     end_date = end_date_entry.get()
#     receive_date = receive_date_entry.get()
#     contact_id = next((k for k, v in suppliers.items() if v == contact_name), None)
#     if not contact_id:
#         messagebox.showerror("Error", "Missing required field: Supplier must be selected.")
#         return
#
#     # Insert the new offer.
#     conn = sqlite3.connect(DB_FILE)
#     cursor = conn.cursor()
#     cursor.execute(
#         "INSERT INTO Offer (offer_end_date, expected_receive_date, contact_id) VALUES (?, ?, ?)",
#         (end_date if end_date else None, receive_date, int(contact_id))
#     )
#     new_offer_id = cursor.lastrowid
#     conn.commit()
#     conn.close()
#     session_offer_ids.append(new_offer_id)
#
#     # Insert rows into Offer_Product and compute investment.
#     investment = 0
#     for product_id, qty_entry in qty_entries.items():
#         if product_vars[product_id].get():
#             try:
#                 quantity = int(qty_entry.get())
#             except ValueError:
#                 messagebox.showerror("Error", "Quantity must be a number!")
#                 return
#             try:
#                 unit_price = float(purchase_price_entries[product_id].get())
#             except ValueError:
#                 messagebox.showerror("Error", "Purchase Unit Price must be numeric!")
#                 return
#             unit = offer_unit_entries[product_id].get().strip()
#             if not unit:
#                 messagebox.showerror("Error", "Unit is required for each selected product!")
#                 return
#             execute_query(
#                 "INSERT INTO Offer_Product (offer_id, product_id, quantity_purchased, purchase_price_per_unit, unit_id) VALUES (?, ?, ?, ?, ?)",
#                 (new_offer_id, product_id, quantity, unit_price, unit)
#             )
#             UpdateDatabaseWithCalculatedFields.update_inventory_by_product_id(product_id)
#
#     #calculated fields for offer_product:
#     #total_product_purchase_price
#     UpdateDatabaseWithCalculatedFields.update_offer_product_by_offer_id(new_offer_id)
#     # For a new offer, there are no customer sales yet.
#
#
#     #calculated fields for offer
#     #expected_flip_date
#     #investment
#     #profit
#     #total_sale_price
#     #sale_complete
#     conn = sqlite3.connect(DB_FILE)
#     cursor = conn.cursor()
#     UpdateDatabaseWithCalculatedFields.update_offer_by_offer_id(new_offer_id)
#     conn.commit()
#     conn.close()
#
#     messagebox.showinfo("Success", f"Offer {new_offer_id} added! Now add Customer Sales.")
#     refresh_dropdowns()
#
#
# def open_add_product():
#     popup = Toplevel(root)
#     popup.title("Add New Product")
#     tk.Label(popup, text="Product Name:").pack()
#     name_entry = tk.Entry(popup)
#     name_entry.pack()
#     tk.Label(popup, text="(Note: Prices for the product are entered with each offer/sale)").pack()
#     tk.Button(popup, text="Add Product", command=lambda: submit_new_product(name_entry, popup)).pack()
#
#
# def submit_new_product(name_entry, popup):
#     name = name_entry.get().strip()
#     if not name:
#         messagebox.showerror("Error", "Product name is required!")
#         return
#     execute_query("INSERT INTO Product (name) VALUES (?)", (name,))
#     messagebox.showinfo("Success", "Product added successfully!")
#     popup.destroy()
#     refresh_dropdowns()
#
#
# def add_contact():
#     name = contact_name_entry.get()
#     contact_type = contact_type_var.get()
#     site = site_entry.get() if site_entry.get() else None
#     phone = phone_entry.get() if phone_entry.get() else None
#     if not name or not contact_type:
#         messagebox.showerror("Error", "Name and Type are required!")
#         return
#     execute_query(
#         "INSERT INTO Contact (name, site, phone, type) VALUES (?, ?, ?, ?)",
#         (name, site, phone, contact_type)
#     )
#     messagebox.showinfo("Success", "Contact added successfully!")
#     refresh_dropdowns()
#     contact_name_entry.delete(0, tk.END)
#     site_entry.delete(0, tk.END)
#     phone_entry.delete(0, tk.END)
#
# #CHANGE majorly
# def add_customer_sale():
#     global offers, customers, products
#     offers = get_offers()
#     customers = get_customers()
#     products = get_products()
#     offer_name = offer_dropdown.get()
#     customer_name = customer_dropdown.get()
#     sell_date = sell_date_entry.get()
#     offer_id = next((k for k, v in offers.items() if v == offer_name), None)
#     if not offer_id:
#         messagebox.showerror("Error", "Offer must be selected.")
#         return
#     if offer_id not in session_offer_ids:
#         session_offer_ids.append(offer_id)
#     customer_id = next((k for k, v in customers.items() if v == customer_name), None)
#     if not customer_id:
#         messagebox.showerror("Error", "Customer must be selected.")
#         return
#
#     selected_products = {}
#     for product_id, var in sale_product_vars.items():
#         if var.get():
#             qty_str = sale_qty_entries[product_id].get()
#             try:
#                 qty = int(qty_str)
#             except ValueError:
#                 messagebox.showerror("Error", "Sale quantity must be a number.")
#                 return
#             price_str = sale_price_entries[product_id].get()
#             try:
#                 unit_sell_price = float(price_str)
#             except ValueError:
#                 messagebox.showerror("Error", "Sell Unit Price must be numeric!")
#                 return
#             sale_unit = sale_unit_entries[product_id].get().strip()
#             if not sale_unit:
#                 messagebox.showerror("Error", "Unit is required for each selected sale product!")
#                 return
#             selected_products[product_id] = (qty, unit_sell_price, sale_unit)
#
#     if not selected_products:
#         messagebox.showerror("Error", "Select at least one product and enter quantity sold, sell price, and unit.")
#         return
#
#     total_sale_price = 0
#     total_qty_sold_converted = 0
#     for product_id, (qty, unit_sell_price, sale_unit) in selected_products.items():
#         purchase_data = execute_query(
#             "SELECT quantity_purchased, unit FROM Offer_Product WHERE offer_id=? AND product_id=?",
#             (offer_id, product_id), fetch=True
#         )
#         if not purchase_data:
#             messagebox.showerror("Error", f"No purchase data found for product {products.get(product_id, '')}.")
#             return
#         purchased_qty, purchase_unit = purchase_data[0]
#         converted_qty = convert_units(qty, sale_unit, purchase_unit)
#         if DEBUG:
#             print('converted qty: ', converted_qty)
#         total_qty_sold_converted += converted_qty
#
#         sales_data = execute_query(
#             """
#             SELECT csp.quantity_sold, u.unit_name AS Sale_Unit
#             FROM CustomerSale_Product csp
#             JOIN CustomerSale cs ON cs.id = csp.customer_sale_id
#             JOIN Unit u ON u.unit_id = csp.unit_id
#             WHERE cs.offer_id=? AND csp.product_id=?
#             """,
#             (offer_id, product_id), fetch=True
#         )
#
#         #CHANGE
#         converted_total = 0
#         for sold_qty, sold_unit in sales_data:
#             converted_total += convert_units(sold_qty, sold_unit, purchase_unit)
#         if DEBUG:
#             print('sum already sold (converted): ', converted_total)
#             print('purchased qty: ', purchased_qty)
#         if converted_qty + converted_total > purchased_qty:
#             left = purchased_qty - converted_total
#             messagebox.showerror("Error",
#                                  f"For product {products.get(product_id, '')}, cannot sell more than {purchased_qty} ({left} left in purchase unit {purchase_unit}).")
#             return
#
#         total_sale_price += qty * unit_sell_price
#
#     conn = sqlite3.connect(DB_FILE)
#     cursor = conn.cursor()
#     today = datetime.date.today()
#     if sell_date:
#         sale_date_obj = datetime.datetime.strptime(sell_date, "%Y-%m-%d").date()
#         sale_complete = sale_date_obj >= today
#     else:
#         sale_complete = False
#
#     cust_phone_data = execute_query("SELECT phone FROM Contact WHERE id=?", (customer_id,), fetch=True)
#     customer_phone = cust_phone_data[0][0] if cust_phone_data else None
#     cursor.execute(
#         "INSERT INTO CustomerSale (sell_date, sale_complete, offer_id, contact_id, quantity_sold, total_sale_price, customer, customer_phone) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
#         (sell_date, sale_complete, offer_id, customer_id, total_qty_sold_converted, total_sale_price, customer_name,
#          customer_phone)
#     )
#     sale_id = cursor.lastrowid
#     conn.commit()
#     conn.close()
#
#     for product_id, (qty, unit_sell_price, sale_unit) in selected_products.items():
#         exists = execute_query(
#             "SELECT COUNT(*) FROM CustomerSale_Product WHERE customer_sale_id=? AND product_id=?",
#             (sale_id, product_id), fetch=True
#         )[0][0]
#         if exists == 0:
#             execute_query(
#                 "INSERT INTO CustomerSale_Product (customer_sale_id, product_id, quantity_sold, sell_price_per_unit, unit) VALUES (?, ?, ?, ?, ?)",
#                 (sale_id, product_id, qty, unit_sell_price, sale_unit)
#             )
#         else:
#             messagebox.showwarning("Warning", f"Product {products.get(product_id, '')} is already linked to this sale; skipping.")
#     session_sale_ids.append(sale_id)
#     messagebox.showinfo("Success", "Customer Sale Added!")
#     UpdateDatabaseWithCalculatedFields.update_customer_sale_product_by_sale_id(sale_id)
#     UpdateDatabaseWithCalculatedFields.update_offer_product_by_offer_id(offer_id)
#     UpdateDatabaseWithCalculatedFields.update_offer_by_offer_id(offer_id)
#     # # ----- Now update the corresponding Offer record -----
#     # conn = sqlite3.connect(DB_FILE)
#     # cursor = conn.cursor()
#     # cursor.execute("SELECT MAX(sell_date) FROM CustomerSale WHERE offer_id = ?", (offer_id,))
#     # expected_flip_date = cursor.fetchone()[0]
#     # print('expected flip date: ', expected_flip_date)
#     # cursor.execute("SELECT COALESCE(SUM(total_sale_price), 0) FROM CustomerSale WHERE offer_id = ?", (offer_id,))
#     # total_sale_price_offer = cursor.fetchone()[0]
#     # print('total sale price offer: ', total_sale_price_offer)
#     # cursor.execute("SELECT COUNT(*) FROM CustomerSale WHERE offer_id = ? AND sale_complete = 0", (offer_id,))
#     # incomplete_sales = cursor.fetchone()[0]
#     # print('incomplete_sales: ', incomplete_sales)
#     # sale_complete_offer = True if incomplete_sales == 0 else False
#     # print('sale complete offer: ', sale_complete_offer)
#     # cursor.execute("SELECT COALESCE(SUM(quantity_purchased * purchase_price_per_unit), 0) FROM Offer_Product WHERE offer_id = ?", (offer_id,))
#     # investment = cursor.fetchone()[0]
#     # profit = total_sale_price_offer - investment if sale_complete_offer else 0
#     # cursor.execute(
#     #     "UPDATE Offer SET expected_flip_date = ?, total_sale_price = ?, sale_complete = ?, profit = ? WHERE id = ?",
#     #     (expected_flip_date, total_sale_price_offer, sale_complete_offer, profit, offer_id)
#     # )
#     # conn.commit()
#     # conn.close()
#     # update_offer_record(offer_id)
#     UpdateDatabaseWithCalculatedFields.update_inventory_by_product_id(product_id)
#     refresh_dropdowns()
#
#
# # def update_offer_record(offer_id):
# #     conn = sqlite3.connect(DB_FILE)
# #     cursor = conn.cursor()
# #
# #     # Get the latest sell_date for this offer.
# #     cursor.execute("SELECT MAX(sell_date) FROM CustomerSale WHERE offer_id = ?", (offer_id,))
# #     expected_flip_date = cursor.fetchone()[0]
# #
# #     # Sum total_sale_price for all CustomerSales for this offer.
# #     cursor.execute("SELECT COALESCE(SUM(total_sale_price), 0) FROM CustomerSale WHERE offer_id = ?", (offer_id,))
# #     total_sale_price_offer = cursor.fetchone()[0]
# #
# #     # Compute the total investment from Offer_Product.
# #     cursor.execute(
# #         "SELECT COALESCE(SUM(quantity_purchased * purchase_price_per_unit), 0) FROM Offer_Product WHERE offer_id = ?",
# #         (offer_id,))
# #     investment = cursor.fetchone()[0]
# #
# #     # Check per product if the sold quantity equals the purchased quantity (with conversion).
# #     cursor.execute("SELECT product_id, quantity_purchased, unit FROM Offer_Product WHERE offer_id = ?", (offer_id,))
# #     offer_products = cursor.fetchall()
# #     offer_complete = True
# #     for prod in offer_products:
# #         prod_id, purchased_qty, purchase_unit = prod
# #         # Sum sold quantities for this product from CustomerSale_Product.
# #         cursor.execute("""
# #             SELECT quantity_sold, unit FROM CustomerSale_Product
# #             WHERE customer_sale_id IN (
# #                 SELECT id FROM CustomerSale WHERE offer_id = ?
# #             ) AND product_id = ?
# #         """, (offer_id, prod_id))
# #         sales = cursor.fetchall()
# #         total_sold = 0
# #         for sale in sales:
# #             qty_sold, sale_unit = sale
# #             total_sold += convert_units(qty_sold, sale_unit, purchase_unit)
# #         if total_sold < purchased_qty:
# #             offer_complete = False
# #             break  # No need to check further if any product is incomplete.
# #
# #     profit = total_sale_price_offer - investment if offer_complete else 0
# #
# #     cursor.execute(
# #         "UPDATE Offer SET expected_flip_date = ?, total_sale_price = ?, sale_complete = ?, profit = ? WHERE id = ?",
# #         (expected_flip_date, total_sale_price_offer, offer_complete, profit, offer_id)
# #     )
# #     conn.commit()
# #     conn.close()
#
#
# # def generate_excel_report():
# #     if not session_offer_ids and not session_sale_ids:
# #         messagebox.showinfo("No Data", "No new offers or sales in this session.")
# #         return
# #
# #     if os.path.exists(REPORT_FILE):
# #         os.remove(REPORT_FILE)
# #         print(f"Deleted old report: {REPORT_FILE}")
# #
# #     if DEBUG:
# #         print("Session Offer IDs:", session_offer_ids)
# #         print("Session Sale IDs:", session_sale_ids)
# #
# #     offer_id_list = ",".join(map(str, session_offer_ids))
# #     sale_id_list = ",".join(map(str, session_sale_ids)) if session_sale_ids else "0"
# #
# #     conn = sqlite3.connect(DB_FILE)
# #     query_offers = f"""
# #     SELECT
# #         o.id AS Offer_ID,
# #         c.name AS SellerName,
# #         c.site AS Supplier_Email,
# #         o.expected_receive_date AS Expected_Receive_Date,
# #         (SELECT MAX(sell_date) FROM CustomerSale WHERE offer_id=o.id AND id IN ({sale_id_list})) AS Expected_Flip_Date,
# #         (
# #             SELECT COALESCE(SUM(op.quantity_purchased * op.purchase_price_per_unit), 0)
# #             FROM Offer_Product op
# #             WHERE op.offer_id=o.id
# #         ) AS Investment,
# #         (
# #             SELECT COUNT(*)
# #             FROM CustomerSale cs
# #             WHERE cs.offer_id=o.id AND cs.id IN ({sale_id_list})
# #         ) AS NumberOfSales
# #     FROM Offer o
# #     JOIN Contact c ON o.contact_id=c.id
# #     WHERE o.id IN ({offer_id_list})
# #     ORDER BY o.id;
# #     """
# #     query_products = f"""
# #     SELECT
# #         o.id AS Offer_ID,
# #         op.product_id AS Product_ID,
# #         p.name AS ProductName,
# #         op.quantity_purchased AS Purchase_Quantity,
# #         op.purchase_price_per_unit AS Purchase_Cost_Per_Unit,
# #         op.unit AS Product_Unit,
# #         (op.quantity_purchased * op.purchase_price_per_unit) AS TotalProductCost
# #     FROM Offer o
# #     JOIN Offer_Product op ON o.id=op.offer_id
# #     JOIN Product p ON op.product_id=p.id
# #     WHERE o.id IN ({offer_id_list})
# #     ORDER BY o.id, p.name;
# #     """
# #     query_sales = f"""
# #     SELECT
# #         o.id AS Offer_ID,
# #         CASE WHEN cs.sale_complete = 1 THEN 'true'
# #              ELSE 'false'
# #         END AS sale_complete,
# #         csp.product_id AS Product_ID,
# #         cs.customer AS CustomerName,
# #         p.name AS Sold_ProductName,
# #         csp.quantity_sold AS Quantity_Sold,
# #         csp.sell_price_per_unit AS SellPricePerUnit,
# #         csp.unit AS Sale_Unit,
# #         (csp.quantity_sold * csp.sell_price_per_unit) AS TotalSalePrice,
# #         cs.sell_date AS Sell_Date
# #     FROM CustomerSale cs
# #     JOIN Offer o ON cs.offer_id=o.id
# #     JOIN CustomerSale_Product csp ON cs.id=csp.customer_sale_id
# #     JOIN Product p ON csp.product_id=p.id
# #     WHERE cs.id IN ({sale_id_list})
# #     ORDER BY o.id, cs.customer;
# #     """
# #
# #     df_offers = pd.read_sql_query(query_offers, conn)
# #     df_products = pd.read_sql_query(query_products, conn)
# #     df_sales = pd.read_sql_query(query_sales, conn)
# #     conn.close()
# #
# #     if DEBUG:
# #         print("----- df_offers -----")
# #         print(df_offers.head())
# #         print("Shape:", df_offers.shape)
# #         print("----- df_products -----")
# #         print(df_products.head())
# #         print("Shape:", df_products.shape)
# #         print("----- df_sales -----")
# #         print(df_sales.head())
# #         print("Shape:", df_sales.shape)
# #
# #     numeric_cols = [
# #         "Investment", "Purchase_Quantity", "Purchase_Cost_Per_Unit", "TotalProductCost",
# #         "Quantity_Sold", "SellPricePerUnit", "TotalSalePrice"
# #     ]
# #     for col in numeric_cols:
# #         if col in df_products.columns:
# #             df_products[col] = pd.to_numeric(df_products[col], errors="coerce").fillna(0)
# #         if col in df_sales.columns:
# #             df_sales[col] = pd.to_numeric(df_sales[col], errors="coerce").fillna(0)
# #         if col in df_offers.columns:
# #             df_offers[col] = pd.to_numeric(df_offers[col], errors="coerce").fillna(0)
# #
# #     if df_sales.empty:
# #         df_offers["TotalOfferFullySold"] = 0
# #     else:
# #         total_offer_fully_sold = df_sales.groupby("Offer_ID")["TotalSalePrice"].sum().reset_index()
# #         total_offer_fully_sold.rename(columns={"TotalSalePrice": "TotalOfferFullySold"}, inplace=True)
# #         df_offers = pd.merge(df_offers, total_offer_fully_sold, on="Offer_ID", how="left")
# #
# #     df_offers["TotalOfferFullySold"] = pd.to_numeric(
# #         df_offers.get("TotalOfferFullySold", 0), errors="coerce"
# #     ).fillna(0)
# #
# #     if not df_sales.empty:
# #         offer_sales_complete = df_sales.groupby("Offer_ID")["sale_complete"].all().reset_index()
# #         offer_sales_complete.rename(columns={"sale_complete": "all_sales_complete"}, inplace=True)
# #         df_offers = pd.merge(df_offers, offer_sales_complete, on="Offer_ID", how="left")
# #     else:
# #         df_offers["all_sales_complete"] = False
# #
# #     df_offers["all_sales_complete"] = pd.array(
# #         np.where(df_offers["all_sales_complete"].isna(), False, df_offers["all_sales_complete"]),
# #         dtype="boolean"
# #     )
# #
# #     df_offers["Profit"] = df_offers.apply(
# #         lambda row: (row["TotalOfferFullySold"] - row["Investment"])
# #         if row.get("all_sales_complete", False) else 0,
# #         axis=1
# #     )
# #
# #     df_offers = df_offers[
# #         [
# #             "Offer_ID", "SellerName", "Supplier_Email", "Expected_Receive_Date", "Expected_Flip_Date",
# #             "Investment", "NumberOfSales", "TotalOfferFullySold", "Profit"
# #         ]
# #     ]
# #
# #     df_products = df_products[
# #         [
# #             "Offer_ID", "Product_ID", "ProductName",
# #             "Purchase_Quantity", "Purchase_Cost_Per_Unit",
# #             "Product_Unit", "TotalProductCost"
# #         ]
# #     ]
# #
# #     df_sales = df_sales[
# #         [
# #             "Offer_ID", "sale_complete", "Product_ID", "CustomerName",
# #             "Sold_ProductName", "Quantity_Sold", "SellPricePerUnit",
# #             "Sale_Unit", "TotalSalePrice", "Sell_Date"
# #         ]
# #     ]
# #
# #     df_merged = pd.merge(
# #         df_products,
# #         df_sales,
# #         on=["Offer_ID", "Product_ID"],
# #         how="outer",
# #         suffixes=("_prod", "_sale")
# #     )
# #
# #     if DEBUG:
# #         print("----- df_merged -----")
# #         print(df_merged.head())
# #         print("Shape:", df_merged.shape)
# #
# #     cascading_data = []
# #     for offer_id in df_offers["Offer_ID"].unique():
# #         off_row = df_offers[df_offers["Offer_ID"] == offer_id].iloc[0]
# #         merged_rows = df_merged[df_merged["Offer_ID"] == offer_id].reset_index(drop=True)
# #         if merged_rows.empty:
# #             cascading_data.append(list(off_row) + [""] * 13)
# #             continue
# #         for i in range(len(merged_rows)):
# #             if i == 0:
# #                 offer_data = list(off_row)
# #             else:
# #                 offer_data = [""] * len(off_row)
# #             product_sale_data = [
# #                 merged_rows.iloc[i]["ProductName"],
# #                 merged_rows.iloc[i]["Purchase_Quantity"],
# #                 merged_rows.iloc[i]["Purchase_Cost_Per_Unit"],
# #                 merged_rows.iloc[i]["Product_Unit"],
# #                 merged_rows.iloc[i]["TotalProductCost"],
# #                 merged_rows.iloc[i]["sale_complete"],
# #                 merged_rows.iloc[i]["CustomerName"],
# #                 merged_rows.iloc[i]["Sold_ProductName"],
# #                 merged_rows.iloc[i]["Quantity_Sold"],
# #                 merged_rows.iloc[i]["SellPricePerUnit"],
# #                 merged_rows.iloc[i]["Sale_Unit"],
# #                 merged_rows.iloc[i]["TotalSalePrice"],
# #                 merged_rows.iloc[i]["Sell_Date"],
# #             ]
# #             row_data = offer_data + product_sale_data
# #             cascading_data.append(row_data)
# #
# #
# #
# #     if DEBUG and cascading_data:
# #         print("----- Cascading Data Sample -----")
# #         for row in cascading_data[:3]:
# #             print(row)
# #         print("Each row length:", len(cascading_data[0]))
# #
# #     final_columns = [
# #         "Offer_ID", "SellerName", "Supplier_Email", "Expected_Receive_Date", "Expected_Flip_Date",
# #         "Investment", "NumberOfSales", "TotalOfferFullySold", "Profit",
# #         "ProductName", "Purchase_Quantity", "Purchase_Cost_Per_Unit", "Product_Unit", "TotalProductCost",
# #         "Sale_Complete", "CustomerName", "Sold_ProductName", "Quantity_Sold", "SellPricePerUnit",
# #         "Sale_Unit", "TotalSalePrice", "Sell_Date"
# #     ]
# #
# #     df_cascading = pd.DataFrame(cascading_data, columns=final_columns)
# #     df_cascading.to_excel(REPORT_FILE, index=False)
# #     messagebox.showinfo("Report Generated", f"New session report created at {REPORT_FILE}")
# #
# #     if DEBUG:
# #         print("Final DataFrame shape:", df_cascading.shape)
# #         print(df_cascading.head())
#
# # ===================== EDIT OFFER FUNCTIONS =====================
#
# def generate_excel_report():
#     if not session_offer_ids and not session_sale_ids:
#         messagebox.showinfo("No Data", "No new offers or sales in this session.")
#         return
#
#     if os.path.exists(REPORT_FILE):
#         os.remove(REPORT_FILE)
#         print(f"Deleted old report: {REPORT_FILE}")
#
#     if DEBUG:
#         print("Session Offer IDs:", session_offer_ids)
#         print("Session Sale IDs:", session_sale_ids)
#
#     offer_id_list = ",".join(map(str, session_offer_ids))
#     sale_id_list = ",".join(map(str, session_sale_ids)) if session_sale_ids else "0"
#
#     conn = sqlite3.connect(DB_FILE)
#
#     query_offers = f"""
#     SELECT
#         o.id AS Offer_ID,
#         c.name AS SellerName,
#         c.site AS Supplier_Email,
#         o.expected_receive_date AS Expected_Receive_Date,
#         o.expected_flip_date AS Expected_Flip_Date,
#         o.investment AS Investment,
#         o.number_of_sales AS NumberOfSales,
#         o.total_sale_price AS TotalOfferFullySold,
#         o.profit AS Profit
#     FROM Offer o
#     JOIN Contact c ON o.contact_id = c.id
#     WHERE o.id IN ({offer_id_list})
#     ORDER BY o.id;
#     """
#
#     query_products = f"""
#     SELECT
#         o.id AS Offer_ID,
#         op.product_id AS Product_ID,
#         p.name AS ProductName,
#         op.quantity_purchased AS Purchase_Quantity,
#         op.purchase_price_per_unit AS Purchase_Cost_Per_Unit,
#         u.unit_name AS Product_Unit,
#         op.total_product_purchase_price AS TotalProductCost
#     FROM Offer o
#     JOIN Offer_Product op ON o.id = op.offer_id
#     JOIN Product p ON op.product_id = p.id
#     JOIN Unit u ON u.unit_id = csp.unit_id
#     WHERE o.id IN ({offer_id_list})
#     ORDER BY o.id, p.name;
#     """
#
#     query_sales = f"""
#     SELECT
#         o.id AS Offer_ID,
#         CASE WHEN cs.sale_complete = 1 THEN 'true'
#              ELSE 'false'
#         END AS sale_complete,
#         csp.product_id AS Product_ID,
#         cs.customer AS CustomerName,
#         p.name AS Sold_ProductName,
#         csp.quantity_sold AS Quantity_Sold,
#         csp.sell_price_per_unit AS SellPricePerUnit,
#         u.unit_name AS Sale_Unit,
#         (csp.quantity_sold * csp.sell_price_per_unit) AS TotalSalePrice,
#         cs.sell_date AS Sell_Date
#     FROM CustomerSale cs
#     JOIN Offer o ON cs.offer_id = o.id
#     JOIN CustomerSale_Product csp ON cs.id = csp.customer_sale_id
#     JOIN Product p ON csp.product_id = p.id
#     JOIN Unit u ON u.unit_id = csp.unit_id
#     WHERE cs.id IN ({sale_id_list})
#     ORDER BY o.id, cs.customer;
#     """
#
#     df_offers = pd.read_sql_query(query_offers, conn)
#     df_products = pd.read_sql_query(query_products, conn)
#     df_sales = pd.read_sql_query(query_sales, conn)
#     conn.close()
#
#     # Optional: convert columns to numeric if needed
#     numeric_cols = [
#         "Investment", "Purchase_Quantity", "Purchase_Cost_Per_Unit", "TotalProductCost",
#         "Quantity_Sold", "SellPricePerUnit", "TotalSalePrice"
#     ]
#     for col in numeric_cols:
#         if col in df_products.columns:
#             df_products[col] = pd.to_numeric(df_products[col], errors="coerce").fillna(0)
#         if col in df_sales.columns:
#             df_sales[col] = pd.to_numeric(df_sales[col], errors="coerce").fillna(0)
#         if col in df_offers.columns:
#             df_offers[col] = pd.to_numeric(df_offers[col], errors="coerce").fillna(0)
#
#     # Since the pre-calculated values are now stored in the database, we can simply use them.
#     # (The merging below is only to produce the cascading report layout.)
#
#     df_merged = pd.merge(df_products, df_sales, on=["Offer_ID", "Product_ID"], how="outer", suffixes=("_prod", "_sale"))
#
#     cascading_data = []
#     for offer_id in df_offers["Offer_ID"].unique():
#         off_row = df_offers[df_offers["Offer_ID"] == offer_id].iloc[0]
#         merged_rows = df_merged[df_merged["Offer_ID"] == offer_id].reset_index(drop=True)
#         if merged_rows.empty:
#             cascading_data.append(list(off_row) + [""] * 13)
#             continue
#         for i in range(len(merged_rows)):
#             if i == 0:
#                 offer_data = list(off_row)
#             else:
#                 offer_data = [""] * len(off_row)
#             product_sale_data = [
#                 merged_rows.iloc[i]["ProductName"],
#                 merged_rows.iloc[i]["Purchase_Quantity"],
#                 merged_rows.iloc[i]["Purchase_Cost_Per_Unit"],
#                 merged_rows.iloc[i]["Product_Unit"],
#                 merged_rows.iloc[i]["TotalProductCost"],
#                 merged_rows.iloc[i]["sale_complete"],
#                 merged_rows.iloc[i]["CustomerName"],
#                 merged_rows.iloc[i]["Sold_ProductName"],
#                 merged_rows.iloc[i]["Quantity_Sold"],
#                 merged_rows.iloc[i]["SellPricePerUnit"],
#                 merged_rows.iloc[i]["Sale_Unit"],
#                 merged_rows.iloc[i]["TotalSalePrice"],
#                 merged_rows.iloc[i]["Sell_Date"],
#             ]
#             row_data = offer_data + product_sale_data
#             cascading_data.append(row_data)
#
#     final_columns = [
#         "Offer_ID", "SellerName", "Supplier_Email", "Expected_Receive_Date", "Expected_Flip_Date",
#         "Investment", "NumberOfSales", "TotalOfferFullySold", "Profit",
#         "ProductName", "Purchase_Quantity", "Purchase_Cost_Per_Unit", "Product_Unit", "TotalProductCost",
#         "Sale_Complete", "CustomerName", "Sold_ProductName", "Quantity_Sold", "SellPricePerUnit",
#         "Sale_Unit", "TotalSalePrice", "Sell_Date"
#     ]
#     df_cascading = pd.DataFrame(cascading_data, columns=final_columns)
#     df_cascading.to_excel(REPORT_FILE, index=False)
#     messagebox.showinfo("Report Generated", f"New session report created at {REPORT_FILE}")
#
#     if DEBUG:
#         print("Final DataFrame shape:", df_cascading.shape)
#         print(df_cascading.head())
#
# def load_offer_details(event):
#     selected_offer_str = edit_offer_dropdown.get()
#     try:
#         # Extract the numeric offer ID (handles multi-digit IDs too)
#         offer_id = int(selected_offer_str[0])
#     except (ValueError, IndexError):
#         print("Offer ID not found for:", selected_offer_str)
#         return
#
#     # Retrieve and update the basic offer fields.
#     conn = sqlite3.connect(DB_FILE)
#     cursor = conn.cursor()
#     cursor.execute(
#         "SELECT offer_end_date, expected_receive_date, contact_id FROM Offer WHERE id = ?",
#         (offer_id,)
#     )
#     result = cursor.fetchone()
#     conn.close()
#     if result:
#         offer_end_date, expected_receive_date, contact_id = result
#         edit_end_date_entry.delete(0, tk.END)
#         edit_end_date_entry.insert(0, offer_end_date if offer_end_date else "")
#         edit_receive_date_entry.delete(0, tk.END)
#         edit_receive_date_entry.insert(0, expected_receive_date if expected_receive_date else "")
#         current_suppliers = get_suppliers()
#         supplier_name = current_suppliers.get(str(contact_id), "")
#         edit_contact_dropdown.set(supplier_name)
#
#     # Retrieve products associated with this offer.
#     conn = sqlite3.connect(DB_FILE)
#     cursor = conn.cursor()
#     cursor.execute(
#         "SELECT product_id, quantity_purchased, purchase_price_per_unit, unit FROM Offer_Product WHERE offer_id = ?",
#         (offer_id,)
#     )
#     rows = cursor.fetchall()
#     conn.close()
#     offer_products = {str(row[0]): (row[1], row[2], row[3]) for row in rows}
#
#     # Clear any preserved manual state to force a full refresh.
#     global edit_offer_state
#     edit_offer_state = {}
#
#     # Rebuild the edit offer product list.
#     refresh_edit_product_list(offer_products)
#
#
# def refresh_edit_product_list(offer_products):
#     global edit_offer_state, edit_qty_entries, edit_product_vars, edit_purchase_price_entries, edit_offer_unit_entries
#     if edit_product_vars:
#         for pid, var in edit_product_vars.items():
#             checked = var.get()
#             qty = edit_qty_entries[pid].get() if pid in edit_qty_entries else ""
#             price = edit_purchase_price_entries[pid].get() if pid in edit_purchase_price_entries else ""
#             unit = edit_offer_unit_entries[pid].get() if pid in edit_offer_unit_entries else ""
#             edit_offer_state[pid] = (checked, qty, price, unit)
#     else:
#         edit_offer_state = {}
#     for widget in edit_product_frame.winfo_children():
#         widget.destroy()
#     edit_qty_entries = {}
#     edit_product_vars = {}
#     edit_purchase_price_entries = {}
#     edit_offer_unit_entries = {}
#
#     current_products = get_products()
#     tk.Label(edit_product_frame, text="Products:").grid(row=0, column=0, sticky="w")
#     tk.Label(edit_product_frame, text="Qty").grid(row=0, column=1)
#     tk.Label(edit_product_frame, text="Unit Price").grid(row=0, column=2)
#     tk.Label(edit_product_frame, text="Unit").grid(row=0, column=3)
#     row_index = 1
#     for product_id, product_name in current_products.items():
#         print('current_products: ', current_products)
#         print('current_products items: ', current_products.items())
#
#         var = tk.BooleanVar()
#         # if product_id in edit_offer_state:
#         #     checked, saved_qty, saved_price, saved_unit = edit_offer_state[product_id]
#         #     var.set(checked)
#         if product_id in offer_products:
#             var.set(True)
#             saved_qty = str(offer_products[product_id][0])
#             saved_price = str(offer_products[product_id][1])
#             saved_unit = offer_products[product_id][2]
#         else:
#             saved_qty = ""
#             saved_price = ""
#             saved_unit = ""
#         chk = tk.Checkbutton(edit_product_frame, text=product_name, variable=var,
#                              command=lambda pid=product_id, v=var: toggle_edit_quantity_field(pid, v))
#         chk.grid(row=row_index, column=0, sticky="w", padx=5, pady=2)
#         qty_entry = tk.Entry(edit_product_frame, width=5)
#         price_entry = tk.Entry(edit_product_frame, width=7)
#         unit_entry = tk.Entry(edit_product_frame, width=5)
#         qty_entry.insert(0, saved_qty)
#         price_entry.insert(0, saved_price)
#         unit_entry.insert(0, saved_unit)
#         if not var.get():
#             qty_entry.config(state="disabled")
#             price_entry.config(state="disabled")
#             unit_entry.config(state="disabled")
#         qty_entry.grid(row=row_index, column=1, padx=5)
#         price_entry.grid(row=row_index, column=2, padx=5)
#         unit_entry.grid(row=row_index, column=3, padx=5)
#         edit_product_vars[product_id] = var
#         edit_qty_entries[product_id] = qty_entry
#         edit_purchase_price_entries[product_id] = price_entry
#         edit_offer_unit_entries[product_id] = unit_entry
#         row_index += 1
#
# def edit_offer():
#     selected_offer_str = edit_offer_dropdown.get()
#     try:
#         offer_id = int(selected_offer_str.split()[0])
#     except (ValueError, IndexError):
#         messagebox.showerror("Error", "No valid offer selected.")
#         return
#     print('offer ID to edit: ', offer_id)
#     current_suppliers = get_suppliers()
#     supplier_name = edit_contact_dropdown.get()
#     contact_id = next((k for k, v in current_suppliers.items() if v == supplier_name), None)
#     if not contact_id:
#         messagebox.showerror("Error", "Supplier must be selected.")
#         return
#     end_date = edit_end_date_entry.get()
#     receive_date = edit_receive_date_entry.get()
#     conn = sqlite3.connect(DB_FILE)
#     cursor = conn.cursor()
#     cursor.execute(
#         "UPDATE Offer SET offer_end_date = ?, expected_receive_date = ?, contact_id = ? WHERE id = ?",
#         (end_date if end_date else None, receive_date if receive_date else None, int(contact_id), offer_id)
#     )
#     conn.commit()
#     # Delete current Offer_Product rows for this offer.
#     cursor.execute("DELETE FROM Offer_Product WHERE offer_id = ?", (offer_id,))
#     for product_id, var in edit_product_vars.items():
#         if var.get():
#             qty_entry = edit_qty_entries[product_id]
#             price_entry = edit_purchase_price_entries[product_id]
#             unit_entry = edit_offer_unit_entries[product_id]
#             try:
#                 quantity = int(qty_entry.get())
#             except ValueError:
#                 messagebox.showerror("Error", "Quantity must be a number!")
#                 conn.close()
#                 return
#             try:
#                 unit_price = float(price_entry.get())
#             except ValueError:
#                 messagebox.showerror("Error", "Purchase Unit Price must be numeric!")
#                 conn.close()
#                 return
#             unit = unit_entry.get().strip()
#             if not unit:
#                 messagebox.showerror("Error", "Unit is required for each selected product!")
#                 conn.close()
#                 return
#             cursor.execute(
#                  "UPDATE Offer_Product SET quantity_purchased = ?, purchase_price_per_unit = ?, unit = ? WHERE offer_id = ? AND product_id = ?",
#                 (quantity, unit_price, unit, offer_id, product_id)
#             )
#             print('offer_product updated: ', offer_id, product_id, quantity, unit_price, unit)
#     conn.commit()
#     conn.close()
#
#     # Now recalculate investment from the Offer_Product table.
#
#     UpdateDatabaseWithCalculatedFields.update_offer_product_by_offer_id(offer_id)
#     UpdateDatabaseWithCalculatedFields.update_offer_by_offer_id(offer_id)
#     UpdateDatabaseWithCalculatedFields.update_inventory_by_product_id(product_id)
#
#     messagebox.showinfo("Success", f"Offer {offer_id} edited successfully!")
#     refresh_dropdowns()
#
# # ===================== MAIN GUI SETUP =====================
# root = tk.Tk()
# root.title("Business Tracker - Database Input")
# root.geometry("600x500")
# notebook = ttk.Notebook(root)
# notebook.pack(pady=10, expand=True)
#
# # ---- Offer Tab ----
# offer_tab = ttk.Frame(notebook)
# notebook.add(offer_tab, text="Add Offer")
# tk.Label(offer_tab, text="Supplier:").pack()
# contact_dropdown = ttk.Combobox(offer_tab, values=list(get_suppliers().values()), state="readonly")
# contact_dropdown.pack()
# tk.Label(offer_tab, text="Offer End Date (YYYY-MM-DD):").pack()
# end_date_entry = tk.Entry(offer_tab)
# end_date_entry.pack()
# tk.Label(offer_tab, text="Expected Receive Date:").pack()
# receive_date_entry = tk.Entry(offer_tab)
# receive_date_entry.pack()
# product_frame = tk.Frame(offer_tab)
# product_frame.pack()
# tk.Button(offer_tab, text="Add New Product", command=open_add_product).pack()
# tk.Button(offer_tab, text="Submit Offer", command=add_offer).pack()
#
# # ---- Edit Offer Tab ----
# edit_offer_tab = ttk.Frame(notebook)
# notebook.add(edit_offer_tab, text="Edit Offer")
# tk.Label(edit_offer_tab, text="Select Offer:").pack()
# edit_offer_dropdown = ttk.Combobox(edit_offer_tab, values=list(get_offers().values()), state="readonly")
# edit_offer_dropdown.pack()
# edit_offer_dropdown.bind("<<ComboboxSelected>>", load_offer_details)
# tk.Label(edit_offer_tab, text="Supplier:").pack()
# edit_contact_dropdown = ttk.Combobox(edit_offer_tab, values=list(get_suppliers().values()), state="readonly")
# edit_contact_dropdown.pack()
# tk.Label(edit_offer_tab, text="Offer End Date (YYYY-MM-DD):").pack()
# edit_end_date_entry = tk.Entry(edit_offer_tab)
# edit_end_date_entry.pack()
# tk.Label(edit_offer_tab, text="Expected Receive Date:").pack()
# edit_receive_date_entry = tk.Entry(edit_offer_tab)
# edit_receive_date_entry.pack()
# tk.Label(edit_offer_tab, text="Products:").pack()
# edit_product_frame = tk.Frame(edit_offer_tab)
# edit_product_frame.pack()
# tk.Button(edit_offer_tab, text="Add New Product", command=open_add_product).pack()
# tk.Button(edit_offer_tab, text="Submit Changes to Offer", command=edit_offer).pack()
#
# # ---- Contact Tab ----
# contact_tab = ttk.Frame(notebook)
# notebook.add(contact_tab, text="Add Contact")
# tk.Label(contact_tab, text="Name:").pack()
# contact_name_entry = tk.Entry(contact_tab)
# contact_name_entry.pack()
# tk.Label(contact_tab, text="Type:").pack()
# contact_type_var = tk.StringVar()
# contact_type_dropdown = ttk.Combobox(contact_tab, textvariable=contact_type_var,
#                                      values=["supplier", "middleman", "customer"], state="readonly")
# contact_type_dropdown.pack()
# tk.Label(contact_tab, text="Site (Optional):").pack()
# site_entry = tk.Entry(contact_tab)
# site_entry.pack()
# tk.Label(contact_tab, text="Phone Number (Optional):").pack()
# phone_entry = tk.Entry(contact_tab)
# phone_entry.pack()
# tk.Button(contact_tab, text="Add Contact", command=add_contact).pack()
#
# # ---- Customer Sale Tab ----
# sale_tab = ttk.Frame(notebook)
# notebook.add(sale_tab, text="Add Customer Sale")
# tk.Label(sale_tab, text="Select Offer:").pack()
# offer_dropdown = ttk.Combobox(sale_tab, values=list(get_offers().values()), state="readonly")
# offer_dropdown.pack()
# offer_dropdown.bind("<<ComboboxSelected>>", refresh_customer_sale_products)
# sale_product_frame = tk.Frame(sale_tab)
# sale_product_frame.pack()
# sale_product_vars = {}
# sale_qty_entries = {}
# sale_price_entries = {}
# sale_unit_entries = {}
#
# def build_sale_product_list():
#     refresh_sale_product_list()
#     pass
#
# def on_close():
#     """Closes the GUI properly and resumes execution in main.py."""
#     root.destroy()  # This closes the Tkinter window
#     root.quit()  # Ensures mainloop stops
#
# offer_dropdown.bind("<<ComboboxSelected>>", lambda e: build_sale_product_list())
# tk.Label(sale_tab, text="Select Customer:").pack()
# customer_dropdown = ttk.Combobox(sale_tab, values=list(get_customers().values()), state="readonly")
# customer_dropdown.pack()
# tk.Label(sale_tab, text="Sell Date (YYYY-MM-DD):").pack()
# sell_date_entry = tk.Entry(sale_tab)
# sell_date_entry.pack()
# tk.Button(sale_tab, text="Submit Customer Sale", command=add_customer_sale).pack()
# tk.Button(sale_tab, text="Generate Session Report", command=generate_excel_report).pack()
# refresh_product_list()
# root.mainloop()
#
# # def on_close():
# #     """Closes the GUI properly and allows execution to continue."""
# #     root.destroy()  # Closes the Tkinter window
# #     root.quit()  # Ensures mainloop() stops cleanly
# # def run():
# #     """Initialize the GUI and start the application."""
# #     refresh_product_list()
# #     root.protocol("WM_DELETE_WINDOW", on_close)  # Ensure proper cleanup
# #     root.mainloop()

import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel
import os
import pandas as pd
import numpy as np
import datetime
import UpdateDatabaseWithCalculatedFields
import sys
import time
import random
import re

print("Using Python:", sys.executable)
print("Python version:", sys.version)
print("sys.path:", sys.path)

import pandas as pd
print("Pandas imported successfully!")

# ===================== CONFIGURATION =====================
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(__file__)
print(base_dir)
DB_FILE = os.path.join(base_dir, "business_tracker.db")
REPORT_FILE = os.path.join(base_dir, "financial_session_report.xlsx")

# ===================== GLOBAL VARIABLES =====================
session_offer_ids = []  # New offer IDs created this session
session_sale_ids = []   # New sale IDs created this session

suppliers = {}
customers = {}
products = {}
offers = {}

# For the Add Offer Tab product widgets:
qty_entries = {}              # product_id -> Entry widget for quantity
product_vars = {}             # product_id -> BooleanVar
purchase_price_entries = {}   # product_id -> Entry widget for purchase price
# Now we store a dict with "var" and "units_dict" (mapping unit_id -> unit_name) for each product.
offer_unit_entries = {}

# For the Customer Sale Tab:
sale_product_vars = {}        # product_id -> BooleanVar
sale_qty_entries = {}         # product_id -> Entry widget for sale quantity
sale_price_entries = {}       # product_id -> Entry widget for sale price
sale_unit_entries = {}        # product_id -> {"var": tk.StringVar, "units_dict": {unit_id: unit_name}}

# For the Edit Offer Tab:
edit_qty_entries = {}         # product_id -> Entry widget for quantity (edit tab)
edit_product_vars = {}        # product_id -> BooleanVar (edit tab)
edit_purchase_price_entries = {}   # product_id -> Entry widget for purchase price (edit tab)
edit_offer_unit_entries = {}       # product_id -> {"var": tk.StringVar, "units_dict": {unit_id: unit_name}}
edit_offer_state = {}              # product_id -> (checked, qty, price, unit_id)

DEBUG = True  # Enable debug prints

# ===================== DATABASE HELPER FUNCTIONS =====================

def execute_query(query, params=(), fetch=False):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(query, params)
    data = cursor.fetchall() if fetch else None
    conn.commit()
    conn.close()
    return data

def generate_excel_report(report_file, session_only=False, session_offer_ids=None, session_sale_ids=None):
    # if not session_offer_ids and not session_sale_ids:
    #     messagebox.showinfo("No Data", "No new offers or sales in this session.")
    #     return

    if os.path.exists(report_file):
        os.remove(report_file)
        print(f"Deleted old report: {report_file}")

    if DEBUG:
        print("Session Offer IDs:", session_offer_ids)
        print("Session Sale IDs:", session_sale_ids)

    if session_only:
        offer_id_list = ",".join(map(str, session_offer_ids))
        sale_id_list = ",".join(map(str, session_sale_ids)) if session_sale_ids else "0"
    else:
        all_offers = execute_query("SELECT id FROM Offer", fetch=True)
        offer_ids = [str(row[0]) for row in all_offers]
        offer_id_list = ",".join(offer_ids)

        all_sales = execute_query("SELECT id FROM CustomerSale", fetch=True)
        sale_ids = [str(row[0]) for row in all_sales]
        sale_id_list = ",".join(sale_ids) if sale_ids else "0"
    conn = sqlite3.connect(DB_FILE)
    query_offers = f"""
    SELECT 
        o.id AS Offer_ID,
        c.name AS SellerName,
        c.site AS Supplier_Email,
        o.expected_receive_date AS Expected_Receive_Date,
        o.expected_flip_date AS Expected_Flip_Date,
        o.investment AS Investment,
        o.number_of_sales AS NumberOfSales,
        o.total_sale_price AS TotalOfferFullySold,
        o.profit AS Profit
    FROM Offer o
    JOIN Contact c ON o.contact_id = c.id
    WHERE o.id IN ({offer_id_list})
    ORDER BY o.id;
    """

    query_products = f"""
    SELECT
        o.id AS Offer_ID,
        op.product_id AS Product_ID,
        p.name AS ProductName,
        op.quantity_purchased AS Purchase_Quantity,
        op.purchase_price_per_unit AS Purchase_Cost_Per_Unit,
        u.unit_name AS Product_Unit,
        op.total_product_purchase_price AS TotalProductCost
    FROM Offer o
    JOIN Offer_Product op ON o.id = op.offer_id
    JOIN Product p ON op.product_id = p.id
    JOIN Unit u ON op.unit_id = u.unit_id
    WHERE o.id IN ({offer_id_list})
    ORDER BY o.id, p.name;
    """

    query_sales = f"""
    SELECT
        o.id AS Offer_ID,
        CASE WHEN cs.sale_complete = 1 THEN 'true'
             ELSE 'false'
        END AS sale_complete,
        csp.product_id AS Product_ID,
        cs.customer AS CustomerName,
        p.name AS Sold_ProductName,
        csp.quantity_sold AS Quantity_Sold,
        csp.sell_price_per_unit AS SellPricePerUnit,
        u.unit_name AS Sale_Unit,
        cs.total_sale_price AS TotalSalePrice,
        cs.sell_date AS Sell_Date
    FROM CustomerSale cs
    JOIN Offer o ON cs.offer_id = o.id
    JOIN CustomerSale_Product csp ON cs.id = csp.customer_sale_id
    JOIN Product p ON csp.product_id = p.id
    JOIN Unit u ON csp.unit_id = u.unit_id
    WHERE cs.id IN ({sale_id_list})
    ORDER BY o.id, cs.customer;
    """

    df_offers = pd.read_sql_query(query_offers, conn)
    df_products = pd.read_sql_query(query_products, conn)
    df_sales = pd.read_sql_query(query_sales, conn)
    conn.close()

    numeric_cols = [
        "Investment", "Purchase_Quantity", "Purchase_Cost_Per_Unit", "TotalProductCost",
        "Quantity_Sold", "SellPricePerUnit", "TotalSalePrice"
    ]
    for col in numeric_cols:
        if col in df_products.columns:
            df_products[col] = pd.to_numeric(df_products[col], errors="coerce").fillna(0)
        if col in df_sales.columns:
            df_sales[col] = pd.to_numeric(df_sales[col], errors="coerce").fillna(0)
        if col in df_offers.columns:
            df_offers[col] = pd.to_numeric(df_offers[col], errors="coerce").fillna(0)

    df_merged = pd.merge(df_products, df_sales, on=["Offer_ID", "Product_ID"], how="outer", suffixes=("_prod", "_sale"))

    cascading_data = []
    for off_id in df_offers["Offer_ID"].unique():
        off_row = df_offers[df_offers["Offer_ID"] == off_id].iloc[0]
        merged_rows = df_merged[df_merged["Offer_ID"] == off_id].reset_index(drop=True)
        if merged_rows.empty:
            cascading_data.append(list(off_row) + [""] * 13)
            continue
        for i in range(len(merged_rows)):
            if i == 0:
                offer_data = list(off_row)
            else:
                offer_data = [""] * len(off_row)
            product_sale_data = [
                merged_rows.iloc[i]["ProductName"],
                merged_rows.iloc[i]["Purchase_Quantity"],
                merged_rows.iloc[i]["Purchase_Cost_Per_Unit"],
                merged_rows.iloc[i]["Product_Unit"],
                merged_rows.iloc[i]["TotalProductCost"],
                merged_rows.iloc[i]["sale_complete"],
                merged_rows.iloc[i]["CustomerName"],
                merged_rows.iloc[i]["Sold_ProductName"],
                merged_rows.iloc[i]["Quantity_Sold"],
                merged_rows.iloc[i]["SellPricePerUnit"],
                merged_rows.iloc[i]["Sale_Unit"],
                merged_rows.iloc[i]["TotalSalePrice"],
                merged_rows.iloc[i]["Sell_Date"],
            ]
            row_data = offer_data + product_sale_data
            cascading_data.append(row_data)

    final_columns = [
        "Offer_ID", "SellerName", "Supplier_Email", "Expected_Receive_Date", "Expected_Flip_Date",
        "Investment", "NumberOfSales", "TotalOfferFullySold", "Profit",
        "ProductName", "Purchase_Quantity", "Purchase_Cost_Per_Unit", "Product_Unit", "TotalProductCost",
        "Sale_Complete", "CustomerName", "Sold_ProductName", "Quantity_Sold", "SellPricePerUnit",
        "Sale_Unit", "TotalSalePrice", "Sell_Date"
    ]
    df_cascading = pd.DataFrame(cascading_data, columns=final_columns)
    df_cascading.to_excel(report_file, sheet_name="Sheet1", index=False)
    messagebox.showinfo("Report Generated", f"New session report created at {report_file}")

    if DEBUG:
        print("Final DataFrame shape:", df_cascading.shape)
        print(df_cascading.head())

    # full report fields
    #     product_id INTEGER PRIMARY KEY,
    #     # product_amount_purchased_all_time REAL DEFAULT 0,    -- In grams
    #     product_amount_in_inventory REAL DEFAULT 0,            -- In grams (purchased minus sold)
    #     # product_inventory_purchase_price REAL DEFAULT 0,       -- Total cost (in currency)
    #     # product_amount_sold_all_time REAL DEFAULT 0,           -- In grams
    #     # product_total_sell_amount REAL DEFAULT 0,              -- Total revenue from sales
    #     # profit_all_time REAL DEFAULT 0,                        -- Overall profit: revenue - cost
    #     # profit_current REAL DEFAULT 0,                         -- Profit on current inventory
    #     # avg_flip_amount REAL DEFAULT 0,                        -- Average sale price per gram
    #     # best_purchase_amount REAL DEFAULT 0,                   -- Best (lowest) purchase price per gram
    #     # best_purchase_date TEXT,
    #     # best_purchase_contact TEXT,
    #     # best_customer_id INTEGER,                              -- Customer (contact id) who purchased the most (in grams)
    #     # best_customer_total_purchased REAL DEFAULT 0,          -- Total grams purchased by that customer
    #     # best_sale_amount REAL DEFAULT 0,                       -- Best (highest) sale price per gram
    #     # best_sale_date TEXT,

# ===================== MAIN GUI SETUP =====================
def setup_gui():
    root = tk.Tk()
    root.title("Business Tracker - Database Input")
    root.geometry("600x500")

    notebook = ttk.Notebook(root)
    notebook.pack(pady=10, expand=True)



    def get_suppliers():
        return {str(c[0]): c[1] for c in execute_query(
            "SELECT id, name FROM Contact WHERE type='supplier'", fetch=True)}

    def get_customers():
        return {str(c[0]): c[1] for c in execute_query(
            "SELECT id, name FROM Contact WHERE type='customer'", fetch=True)}

    def get_products():
        return {str(p[0]): p[1] for p in execute_query(
            "SELECT id, name FROM Product", fetch=True)}

    def get_offers():
        rows = execute_query(
            """
            SELECT
                o.id,
                c.name,
                Investment
            FROM Offer o
            JOIN Contact c ON o.contact_id = c.id
            """,
            fetch=True
        )
        # Build strings like "2 - SoCal ($1776.0)"; keys kept as numbers.
        return {row[0]: f"{row[0]} - {row[1]} (${row[2]})" for row in rows}

    def get_products_by_offer(offer_id):
        results = execute_query(
            """
            SELECT p.id, p.name, op.quantity_purchased
            FROM Offer_Product op
            JOIN Product p ON op.product_id = p.id
            WHERE op.offer_id = ?
            """, (int(offer_id),), fetch=True)
        return {r[0]: (r[1], r[2]) for r in results}

    # ===================== UNIT HELPER FUNCTIONS =====================
    def get_units():
        """
        Returns a dict mapping unit_id (as string) to unit_name for all units in the Unit table.
        """
        rows = execute_query(
            "SELECT unit_id, unit_name FROM Unit",
            fetch=True
        )
        return {str(r[0]): r[1] for r in rows}

    # ===================== REFRESH FUNCTIONS =====================
    def refresh_dropdowns():
        global suppliers, customers, products, offers
        UpdateDatabaseWithCalculatedFields.update_database()
        suppliers = get_suppliers()
        customers = get_customers()
        products = get_products()
        offers = get_offers()
        if DEBUG:
            print('refreshed dropdowns')
            print('offers:', offers)
        contact_dropdown["values"] = list(suppliers.values())
        customer_dropdown["values"] = list(customers.values())
        offer_dropdown["values"] = list(offers.values())
        try:
            edit_offer_dropdown["values"] = list(offers.values())
            edit_contact_dropdown["values"] = list(suppliers.values())
        except Exception:
            pass
        refresh_product_list()
        refresh_sale_product_list()
        selected = edit_offer_dropdown.get().strip()
        if selected:
            load_offer_details(None)

    def refresh_product_list():
        """
        For the Add Offer tab:
        Lists all products with checkboxes, quantity, unit price,
        and a combobox for selecting the unit (populated from the Unit table globally).
        """
        global products, qty_entries, product_vars, purchase_price_entries, offer_unit_entries
        products = get_products()
        for widget in product_frame.winfo_children():
            widget.destroy()
        qty_entries = {}
        product_vars = {}
        purchase_price_entries = {}
        offer_unit_entries = {}

        tk.Label(product_frame, text="Products:").grid(row=0, column=0, sticky="w")
        tk.Label(product_frame, text="Qty").grid(row=0, column=1)
        tk.Label(product_frame, text="Unit Price").grid(row=0, column=2)
        tk.Label(product_frame, text="Unit").grid(row=0, column=3)

        # Use all available units (global)
        global_units = get_units()

        row_index = 1
        for product_id, product_name in products.items():
            var = tk.BooleanVar()
            chk = tk.Checkbutton(product_frame, text=product_name, variable=var,
                                 command=lambda pid=product_id, v=var: toggle_quantity_field(pid, v))
            chk.grid(row=row_index, column=0, sticky="w", padx=5, pady=2)

            qty_entry = tk.Entry(product_frame, state="disabled", width=5)
            qty_entry.grid(row=row_index, column=1, padx=5)

            price_entry = tk.Entry(product_frame, state="disabled", width=7)
            price_entry.grid(row=row_index, column=2, padx=5)

            # Create a combobox for unit selection
            unit_var = tk.StringVar()
            unit_dropdown = ttk.Combobox(product_frame, textvariable=unit_var, state="disabled", width=7)
            unit_dropdown["values"] = list(global_units.values())
            unit_dropdown.grid(row=row_index, column=3, padx=5)

            product_vars[product_id] = var
            qty_entries[product_id] = qty_entry
            purchase_price_entries[product_id] = price_entry
            # For global units, we store the same units dict for all products.
            offer_unit_entries[product_id] = {"var": unit_var, "widget": unit_dropdown, "units_dict": global_units}

            row_index += 1

    def refresh_sale_product_list():
        try:
            selected_offer_str = offer_dropdown.get()
            offer_id = int(selected_offer_str.split()[0])
        except (ValueError, IndexError) as e:
            if DEBUG:
                print("refresh_sale_product_list: No valid offer selected:", e)
            return

        for widget in sale_product_frame.winfo_children():
            widget.destroy()
        sale_product_vars.clear()
        sale_qty_entries.clear()
        sale_price_entries.clear()
        sale_unit_entries.clear()

        if not offer_id:
            return

        offer_products = get_products_by_offer(offer_id)
        tk.Label(sale_product_frame, text="Products for Sale:").grid(row=0, column=0, sticky="w")
        tk.Label(sale_product_frame, text="Qty").grid(row=0, column=1)
        tk.Label(sale_product_frame, text="Unit Sell Price").grid(row=0, column=2)
        tk.Label(sale_product_frame, text="Unit").grid(row=0, column=3)

        # Use global units for sale as well.
        global_units = get_units()

        row_index = 1
        for product_id, (product_name, _) in offer_products.items():
            var = tk.BooleanVar()
            chk = tk.Checkbutton(sale_product_frame, text=product_name, variable=var)
            chk.grid(row=row_index, column=0, sticky="w", padx=5, pady=2)

            qty_entry = tk.Entry(sale_product_frame, state="disabled", width=5)
            qty_entry.grid(row=row_index, column=1, padx=5)

            price_entry = tk.Entry(sale_product_frame, state="disabled", width=7)
            price_entry.grid(row=row_index, column=2, padx=5)

            unit_var = tk.StringVar()
            unit_dropdown = ttk.Combobox(sale_product_frame, textvariable=unit_var, state="disabled", width=7)
            unit_dropdown["values"] = list(global_units.values())
            unit_dropdown.grid(row=row_index, column=3, padx=5)

            def on_toggle(pid=product_id, v=var, q_entry=qty_entry, p_entry=price_entry, combobox=unit_dropdown):
                if v.get():
                    q_entry.config(state="normal")
                    p_entry.config(state="normal")
                    combobox.config(state="readonly")
                else:
                    q_entry.config(state="disabled")
                    p_entry.config(state="disabled")
                    combobox.config(state="disabled")

            chk.config(command=on_toggle)
            sale_product_vars[product_id] = var
            sale_qty_entries[product_id] = qty_entry
            sale_price_entries[product_id] = price_entry
            sale_unit_entries[product_id] = {"var": unit_var, "units_dict": global_units}
            row_index += 1

    def build_sale_product_list():
        refresh_sale_product_list()

    def refresh_customer_sale_products(event=None):
        refresh_sale_product_list()

    def toggle_quantity_field(product_id, var):
        if var.get():
            qty_entries[product_id].config(state="normal")
            purchase_price_entries[product_id].config(state="normal")
            offer_unit_entries[product_id]["var"].set("")  # Reset unit selection

            # Get the actual dropdown widget and enable it
            unit_dropdown_widget = offer_unit_entries[product_id]["widget"]
            unit_dropdown_widget.config(state="readonly")  # ✅ Correct way to enable dropdown
        else:
            qty_entries[product_id].config(state="disabled")
            purchase_price_entries[product_id].config(state="disabled")

            # Get the actual dropdown widget and disable it
            unit_dropdown_widget = offer_unit_entries[product_id]["widget"]
            unit_dropdown_widget.config(state="disabled")  # ✅ Correct way to disable dropdown

    def toggle_edit_quantity_field(product_id, var):
        if var.get():
            edit_qty_entries[product_id].config(state="normal")
            edit_purchase_price_entries[product_id].config(state="normal")
        else:
            edit_qty_entries[product_id].config(state="disabled")
            edit_purchase_price_entries[product_id].config(state="disabled")

    # ===================== ADD UNIT BUTTON + LOGIC =====================
    def open_add_unit():
        """
        Opens a popup to add a new unit.
        Since Unit table is now global, we do not select a product.
        The user enters:
          - New Unit Name
          - Related Unit Name
          - Conversion Factor (for new_unit → related_unit)
        Two records are inserted: one for forward conversion, one for reverse.
        """
        popup = Toplevel(root)
        popup.title("Add New Unit")

        tk.Label(popup, text="New Unit Name:").pack()
        new_unit_name_entry = tk.Entry(popup)
        new_unit_name_entry.pack()

        tk.Label(popup, text="Related Unit Name:").pack()
        related_unit_name_entry = tk.Entry(popup)
        related_unit_name_entry.pack()

        tk.Label(popup, text="Conversion Factor (new_unit → related_unit):").pack()
        conversion_factor_entry = tk.Entry(popup)
        conversion_factor_entry.pack()

        def submit_new_unit():
            new_unit_name = new_unit_name_entry.get().strip()
            related_unit_name = related_unit_name_entry.get().strip()
            try:
                conv_factor = float(conversion_factor_entry.get())
            except ValueError:
                messagebox.showerror("Error", "Conversion factor must be numeric.")
                return

            if not new_unit_name or not related_unit_name or conv_factor <= 0:
                messagebox.showerror("Error", "All fields are required and factor must be > 0.")
                return

            # Insert the forward record
            execute_query(
                "INSERT INTO Unit (unit_name, conversion_factor) VALUES (?, ?)",
                (new_unit_name, conv_factor)
            )
            new_unit_id = execute_query("SELECT last_insert_rowid()", fetch=True)[0][0]

            # Insert the reverse record with reciprocal conversion factor
            reciprocal = 1.0 / conv_factor
            execute_query(
                "INSERT INTO Unit (unit_name, conversion_factor) VALUES (?, ?)",
                (related_unit_name, reciprocal)
            )
            reverse_unit_id = execute_query("SELECT last_insert_rowid()", fetch=True)[0][0]

            # Link them as related units
            execute_query("UPDATE Unit SET related_unit_id = ? WHERE unit_id = ?", (reverse_unit_id, new_unit_id))
            execute_query("UPDATE Unit SET related_unit_id = ? WHERE unit_id = ?", (new_unit_id, reverse_unit_id))

            messagebox.showinfo("Success", f"Added units '{new_unit_name}' and '{related_unit_name}'.")
            popup.destroy()
            refresh_dropdowns()

        tk.Button(popup, text="Add Unit", command=submit_new_unit).pack()

    # ===================== ADD/UPDATE FUNCTIONS =====================
    def open_add_product():
        popup = Toplevel(root)
        popup.title("Add New Product")
        tk.Label(popup, text="Product Name:").pack()
        name_entry = tk.Entry(popup)
        name_entry.pack()
        tk.Label(popup, text="(Note: Prices for the product are entered with each offer/sale)").pack()
        tk.Button(popup, text="Add Product", command=lambda: submit_new_product(name_entry, popup)).pack()

    def submit_new_product(name_entry, popup):
        name = name_entry.get().strip()
        if not name:
            messagebox.showerror("Error", "Product name is required!")
            return
        execute_query("INSERT INTO Product (name) VALUES (?)", (name,))
        messagebox.showinfo("Success", "Product added successfully!")
        popup.destroy()
        refresh_dropdowns()

    def add_contact():
        name = contact_name_entry.get()
        contact_type = contact_type_var.get()
        site = site_entry.get() if site_entry.get() else None
        phone = phone_entry.get() if phone_entry.get() else None
        if not name or not contact_type:
            messagebox.showerror("Error", "Name and Type are required!")
            return
        execute_query(
            "INSERT INTO Contact (name, site, phone, type) VALUES (?, ?, ?, ?)",
            (name, site, phone, contact_type)
        )
        messagebox.showinfo("Success", "Contact added successfully!")
        refresh_dropdowns()
        contact_name_entry.delete(0, tk.END)
        site_entry.delete(0, tk.END)
        phone_entry.delete(0, tk.END)

    def add_offer():
        global suppliers
        suppliers = get_suppliers()
        contact_name = contact_dropdown.get()
        end_date = end_date_entry.get()
        receive_date = receive_date_entry.get()
        contact_id = next((k for k, v in suppliers.items() if v == contact_name), None)
        if not contact_id:
            messagebox.showerror("Error", "Missing required field: Supplier must be selected.")
            return

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO Offer (offer_end_date, expected_receive_date, contact_id) VALUES (?, ?, ?)",
            (end_date if end_date else None, receive_date, int(contact_id))
        )
        new_offer_id = cursor.lastrowid
        conn.commit()
        conn.close()
        session_offer_ids.append(new_offer_id)

        # Insert rows into Offer_Product for each selected product
        for product_id, qty_entry in qty_entries.items():
            if product_vars[product_id].get():
                try:
                    quantity = int(qty_entry.get())
                except ValueError:
                    messagebox.showerror("Error", "Quantity must be a number!")
                    return
                try:
                    unit_price = float(purchase_price_entries[product_id].get())
                except ValueError:
                    messagebox.showerror("Error", "Purchase Unit Price must be numeric!")
                    return
                combo_info = offer_unit_entries[product_id]
                unit_var = combo_info["var"]
                units_dict = combo_info["units_dict"]
                unit_name = unit_var.get().strip()
                if not unit_name:
                    messagebox.showerror("Error", "Unit is required for each selected product!")
                    return
                unit_id_str = next((k for k, v in units_dict.items() if v == unit_name), None)
                if not unit_id_str:
                    messagebox.showerror("Error", f"Could not find unit_id for unit name '{unit_name}'.")
                    return
                unit_id = int(unit_id_str)
                execute_query(
                    "INSERT INTO Offer_Product (offer_id, product_id, quantity_purchased, purchase_price_per_unit, unit_id) VALUES (?, ?, ?, ?, ?)",
                    (new_offer_id, product_id, quantity, unit_price, unit_id)
                )
        messagebox.showinfo("Success", f"Offer {new_offer_id} added! Now add Customer Sales.")
        refresh_dropdowns()

    def add_customer_sale():
        global offers, customers, products
        offers = get_offers()
        customers = get_customers()
        products = get_products()
        offer_name = offer_dropdown.get()
        customer_name = customer_dropdown.get()
        sell_date = sell_date_entry.get()
        offer_id = next((k for k, v in offers.items() if v == offer_name), None)
        if not offer_id:
            messagebox.showerror("Error", "Offer must be selected.")
            return
        if offer_id not in session_offer_ids:
            session_offer_ids.append(offer_id)
        customer_id = next((k for k, v in customers.items() if v == customer_name), None)
        if not customer_id:
            messagebox.showerror("Error", "Customer must be selected.")
            return

        selected_products = {}
        for product_id, var in sale_product_vars.items():
            if var.get():
                qty_str = sale_qty_entries[product_id].get()
                try:
                    qty = int(qty_str)
                except ValueError:
                    messagebox.showerror("Error", "Sale quantity must be a number.")
                    return
                price_str = sale_price_entries[product_id].get()
                try:
                    unit_sell_price = float(price_str)
                except ValueError:
                    messagebox.showerror("Error", "Sell Unit Price must be numeric!")
                    return
                combo_info = sale_unit_entries[product_id]
                unit_var = combo_info["var"]
                units_dict = combo_info["units_dict"]
                unit_name = unit_var.get().strip()
                if not unit_name:
                    messagebox.showerror("Error", "Unit is required for each selected sale product!")
                    return
                unit_id_str = next((k for k, v in units_dict.items() if v == unit_name), None)
                if not unit_id_str:
                    messagebox.showerror("Error", f"Could not find unit_id for unit name '{unit_name}'.")
                    return
                unit_id = int(unit_id_str)
                selected_products[product_id] = (qty, unit_sell_price, unit_id)

        if not selected_products:
            messagebox.showerror("Error", "Select at least one product and enter quantity sold, sell price, and unit.")
            return

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cust_phone_data = execute_query("SELECT phone FROM Contact WHERE id=?", (customer_id,), fetch=True)
        customer_phone = cust_phone_data[0][0] if cust_phone_data else None
        cursor.execute(
            "INSERT INTO CustomerSale (sell_date, offer_id, contact_id, customer, customer_phone) VALUES (?, ?, ?, ?, ?)",
            (sell_date, offer_id, customer_id, customer_name, customer_phone)
        )
        sale_id = cursor.lastrowid
        conn.commit()
        conn.close()

        for product_id, (qty, unit_sell_price, unit_id) in selected_products.items():
            exists = execute_query(
                "SELECT COUNT(*) FROM CustomerSale_Product WHERE customer_sale_id=? AND product_id=?",
                (sale_id, product_id), fetch=True
            )[0][0]
            if exists == 0:
                execute_query(
                    "INSERT INTO CustomerSale_Product (customer_sale_id, product_id, quantity_sold, sell_price_per_unit, unit_id) VALUES (?, ?, ?, ?, ?)",
                    (sale_id, product_id, qty, unit_sell_price, unit_id)
                )
            else:
                messagebox.showwarning("Warning",
                                       f"Product {products.get(product_id, '')} is already linked to this sale; skipping.")
        session_sale_ids.append(sale_id)
        messagebox.showinfo("Success", "Customer Sale Added!")
        refresh_dropdowns()

    def load_offer_details(event):
        selected_offer_str = edit_offer_dropdown.get()
        try:
            offer_id = int(selected_offer_str.split()[0])
        except (ValueError, IndexError):
            print("Offer ID not found for:", selected_offer_str)
            return

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT offer_end_date, expected_receive_date, contact_id FROM Offer WHERE id = ?", (offer_id,))
        result = cursor.fetchone()
        conn.close()
        if result:
            offer_end_date, expected_receive_date, contact_id = result
            edit_end_date_entry.delete(0, tk.END)
            edit_end_date_entry.insert(0, offer_end_date if offer_end_date else "")
            edit_receive_date_entry.delete(0, tk.END)
            edit_receive_date_entry.insert(0, expected_receive_date if expected_receive_date else "")
            current_suppliers = get_suppliers()
            supplier_name = current_suppliers.get(str(contact_id), "")
            edit_contact_dropdown.set(supplier_name)

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT product_id, quantity_purchased, purchase_price_per_unit, unit_id FROM Offer_Product WHERE offer_id = ?",
            (offer_id,)
        )
        rows = cursor.fetchall()
        conn.close()

        offer_products = {}
        for r in rows:
            prod_id, qty, price, unit_id = r
            offer_products[str(prod_id)] = (qty, price, unit_id)

        global edit_offer_state
        edit_offer_state = {}
        refresh_edit_product_list(offer_products)

    def refresh_edit_product_list(offer_products):
        global edit_offer_state, edit_qty_entries, edit_product_vars, edit_purchase_price_entries, edit_offer_unit_entries
        if edit_product_vars:
            for pid, var in edit_product_vars.items():
                checked = var.get()
                qty = edit_qty_entries[pid].get() if pid in edit_qty_entries else ""
                price = edit_purchase_price_entries[pid].get() if pid in edit_purchase_price_entries else ""
                edit_offer_state[pid] = (checked, qty, price)
        else:
            edit_offer_state = {}
        for widget in edit_product_frame.winfo_children():
            widget.destroy()
        edit_qty_entries = {}
        edit_product_vars = {}
        edit_purchase_price_entries = {}
        edit_offer_unit_entries = {}

        current_products = get_products()
        tk.Label(edit_product_frame, text="Products:").grid(row=0, column=0, sticky="w")
        tk.Label(edit_product_frame, text="Qty").grid(row=0, column=1)
        tk.Label(edit_product_frame, text="Unit Price").grid(row=0, column=2)
        tk.Label(edit_product_frame, text="Unit").grid(row=0, column=3)
        row_index = 1
        global_units = get_units()  # All units globally
        for product_id, product_name in current_products.items():
            var = tk.BooleanVar()
            saved_qty = ""
            saved_price = ""
            saved_unit_id = None
            if product_id in offer_products:
                var.set(True)
                saved_qty = str(offer_products[product_id][0])
                saved_price = str(offer_products[product_id][1])
                saved_unit_id = offer_products[product_id][2]

            chk = tk.Checkbutton(edit_product_frame, text=product_name, variable=var,
                                 command=lambda pid=product_id, v=var: toggle_edit_quantity_field(pid, v))
            chk.grid(row=row_index, column=0, sticky="w", padx=5, pady=2)

            qty_entry = tk.Entry(edit_product_frame, width=5)
            price_entry = tk.Entry(edit_product_frame, width=7)
            qty_entry.insert(0, saved_qty)
            price_entry.insert(0, saved_price)
            if not var.get():
                qty_entry.config(state="disabled")
                price_entry.config(state="disabled")
            qty_entry.grid(row=row_index, column=1, padx=5)
            price_entry.grid(row=row_index, column=2, padx=5)

            unit_var = tk.StringVar()
            unit_dropdown = ttk.Combobox(edit_product_frame, textvariable=unit_var, state="disabled", width=7)
            unit_dropdown["values"] = list(global_units.values())
            if saved_unit_id and str(saved_unit_id) in global_units:
                unit_var.set(global_units[str(saved_unit_id)])
            if var.get():
                unit_dropdown.config(state="readonly")
            unit_dropdown.grid(row=row_index, column=3, padx=5)

            edit_product_vars[product_id] = var
            edit_qty_entries[product_id] = qty_entry
            edit_purchase_price_entries[product_id] = price_entry
            edit_offer_unit_entries[product_id] = {"var": unit_var, "units_dict": global_units}
            row_index += 1

    def edit_offer():
        selected_offer_str = edit_offer_dropdown.get()
        try:
            offer_id = int(selected_offer_str.split()[0])
        except (ValueError, IndexError):
            messagebox.showerror("Error", "No valid offer selected.")

        current_suppliers = get_suppliers()
        supplier_name = edit_contact_dropdown.get()
        contact_id = next((k for k, v in current_suppliers.items() if v == supplier_name), None)
        if not contact_id:
            messagebox.showerror("Error", "Supplier must be selected.")
            return

        end_date = edit_end_date_entry.get()
        receive_date = edit_receive_date_entry.get()

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE Offer SET offer_end_date = ?, expected_receive_date = ?, contact_id = ? WHERE id = ?",
            (end_date if end_date else None, receive_date if receive_date else None, int(contact_id), offer_id)
        )
        conn.commit()

        cursor.execute("DELETE FROM Offer_Product WHERE offer_id = ?", (offer_id,))
        conn.commit()

        for product_id, var in edit_product_vars.items():
            if var.get():
                qty_entry = edit_qty_entries[product_id]
                price_entry = edit_purchase_price_entries[product_id]
                combo_info = edit_offer_unit_entries[product_id]
                unit_var = combo_info["var"]
                units_dict = combo_info["units_dict"]

                try:
                    quantity = int(qty_entry.get())
                except ValueError:
                    messagebox.showerror("Error", "Quantity must be a number!")
                    conn.close()
                    return
                try:
                    unit_price = float(price_entry.get())
                except ValueError:
                    messagebox.showerror("Error", "Purchase Unit Price must be numeric!")
                    conn.close()
                    return

                unit_name = unit_var.get().strip()
                if not unit_name:
                    messagebox.showerror("Error", "Unit is required for each selected product!")
                    conn.close()
                    return

                unit_id_str = next((k for k, v in units_dict.items() if v == unit_name), None)
                if not unit_id_str:
                    messagebox.showerror("Error", f"Could not find unit_id for unit name '{unit_name}'.")
                    conn.close()
                    return
                unit_id = int(unit_id_str)
                cursor.execute(
                    "INSERT INTO Offer_Product (offer_id, product_id, quantity_purchased, purchase_price_per_unit, unit_id) VALUES (?, ?, ?, ?, ?)",
                    (offer_id, product_id, quantity, unit_price, unit_id)
                )
        conn.commit()
        conn.close()

        messagebox.showinfo("Success", f"Offer {offer_id} edited successfully!")
        refresh_dropdowns()

    # ---- Offer Tab ----
    offer_tab = ttk.Frame(notebook)
    notebook.add(offer_tab, text="Add Offer")
    tk.Label(offer_tab, text="Supplier:").pack()
    contact_dropdown = ttk.Combobox(offer_tab, values=list(get_suppliers().values()), state="readonly")
    contact_dropdown.pack()
    tk.Label(offer_tab, text="Offer End Date (YYYY-MM-DD):").pack()
    end_date_entry = tk.Entry(offer_tab)
    end_date_entry.pack()
    tk.Label(offer_tab, text="Expected Receive Date:").pack()
    receive_date_entry = tk.Entry(offer_tab)
    receive_date_entry.pack()

    product_frame = tk.Frame(offer_tab)
    product_frame.pack()

    tk.Button(offer_tab, text="Add New Product", command=open_add_product).pack()
    tk.Button(offer_tab, text="Add Unit", command=open_add_unit).pack()
    tk.Button(offer_tab, text="Submit Offer", command=add_offer).pack()

    # ---- Edit Offer Tab ----
    edit_offer_tab = ttk.Frame(notebook)
    notebook.add(edit_offer_tab, text="Edit Offer")
    tk.Label(edit_offer_tab, text="Select Offer:").pack()
    edit_offer_dropdown = ttk.Combobox(edit_offer_tab, values=list(get_offers().values()), state="readonly")
    edit_offer_dropdown.pack()
    edit_offer_dropdown.bind("<<ComboboxSelected>>", load_offer_details)
    tk.Label(edit_offer_tab, text="Supplier:").pack()
    edit_contact_dropdown = ttk.Combobox(edit_offer_tab, values=list(get_suppliers().values()), state="readonly")
    edit_contact_dropdown.pack()
    tk.Label(edit_offer_tab, text="Offer End Date (YYYY-MM-DD):").pack()
    edit_end_date_entry = tk.Entry(edit_offer_tab)
    edit_end_date_entry.pack()
    tk.Label(edit_offer_tab, text="Expected Receive Date:").pack()
    edit_receive_date_entry = tk.Entry(edit_offer_tab)
    edit_receive_date_entry.pack()

    tk.Label(edit_offer_tab, text="Products:").pack()
    edit_product_frame = tk.Frame(edit_offer_tab)
    edit_product_frame.pack()

    tk.Button(edit_offer_tab, text="Add New Product", command=open_add_product).pack()
    tk.Button(edit_offer_tab, text="Add Unit", command=open_add_unit).pack()
    tk.Button(edit_offer_tab, text="Submit Changes to Offer", command=edit_offer).pack()

    # ---- Contact Tab ----
    contact_tab = ttk.Frame(notebook)
    notebook.add(contact_tab, text="Add Contact")
    tk.Label(contact_tab, text="Name:").pack()
    contact_name_entry = tk.Entry(contact_tab)
    contact_name_entry.pack()
    tk.Label(contact_tab, text="Type:").pack()
    contact_type_var = tk.StringVar()
    contact_type_dropdown = ttk.Combobox(contact_tab, textvariable=contact_type_var,
                                         values=["supplier", "middleman", "customer"], state="readonly")
    contact_type_dropdown.pack()
    tk.Label(contact_tab, text="Site (Optional):").pack()
    site_entry = tk.Entry(contact_tab)
    site_entry.pack()
    tk.Label(contact_tab, text="Phone Number (Optional):").pack()
    phone_entry = tk.Entry(contact_tab)
    phone_entry.pack()
    tk.Button(contact_tab, text="Add Contact", command=add_contact).pack()

    # ---- Customer Sale Tab ----
    sale_tab = ttk.Frame(notebook)
    notebook.add(sale_tab, text="Add Customer Sale")
    tk.Label(sale_tab, text="Select Offer:").pack()
    offer_dropdown = ttk.Combobox(sale_tab, values=list(get_offers().values()), state="readonly")
    offer_dropdown.pack()
    offer_dropdown.bind("<<ComboboxSelected>>", refresh_customer_sale_products)

    sale_product_frame = tk.Frame(sale_tab)
    sale_product_frame.pack()

    tk.Button(sale_tab, text="Add Unit", command=open_add_unit).pack()

    tk.Label(sale_tab, text="Select Customer:").pack()
    customer_dropdown = ttk.Combobox(sale_tab, values=list(get_customers().values()), state="readonly")
    customer_dropdown.pack()

    tk.Label(sale_tab, text="Sell Date (YYYY-MM-DD):").pack()
    sell_date_entry = tk.Entry(sale_tab)
    sell_date_entry.pack()

    tk.Button(sale_tab, text="Submit Customer Sale", command=add_customer_sale).pack()
    tk.Button(sale_tab, text="Generate Session Report", command=lambda: generate_excel_report(REPORT_FILE, True, session_offer_ids=session_offer_ids, session_sale_ids=session_sale_ids)).pack()



    refresh_product_list()
    root.mainloop()

if __name__ == "__main__":
    setup_gui()