import os
import sqlite3
import datetime as stdlib_datetime
import sys
from datetime import date



if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(__file__)
print(base_dir)
DB_FILE = os.path.join(base_dir, "business_tracker.db")



def update_customer_sale():
    # Get today's date in YYYY-MM-DD format
    # today = stdlib_datetime.date.today()
    today = date.today().strftime("%Y-%m-%d")
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    # Update the table to mark sales as complete
    cursor.execute(""" 
        UPDATE CustomerSale 
        SET sale_complete = 1,
        total_sale_price = (
            SELECT COALESCE(SUM(csp.total_product_sale_price), 0)
            FROM CustomerSale_Product csp
            WHERE CustomerSale.id = csp.customer_sale_id
        )
        WHERE sell_date <= ?
    """,(today,))
    conn.commit()
    conn.close()


def update_offer():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
    UPDATE Offer
    SET 
        expected_flip_date = (
            SELECT MAX(sell_date)
            FROM CustomerSale 
            WHERE offer_id = Offer.id
        ),
        total_sale_price = (
            SELECT COALESCE(SUM(total_sale_price), 0)
            FROM CustomerSale 
            WHERE offer_id = Offer.id
        ),
        sale_complete = CASE 
                            WHEN (SELECT COUNT(*) 
                                  FROM CustomerSale 
                                  WHERE offer_id = Offer.id 
                                    AND sale_complete = 0) = 0 
                            THEN 1 
                            ELSE 0 
                        END,
        investment = (
            SELECT COALESCE(SUM(total_product_purchase_price), 0)
            FROM Offer_Product 
            WHERE offer_id = Offer.id
        ),
        profit = CASE 
                    WHEN (SELECT COUNT(*) 
                          FROM CustomerSale 
                          WHERE offer_id = Offer.id 
                            AND sale_complete = 0) = 0 
                    THEN (
                        (SELECT COALESCE(SUM(total_sale_price), 0)
                         FROM CustomerSale 
                         WHERE offer_id = Offer.id)
                        -
                        (SELECT COALESCE(SUM(quantity_purchased * purchase_price_per_unit), 0)
                         FROM Offer_Product 
                         WHERE offer_id = Offer.id)
                    )
                    ELSE 0 
                 END,
        number_of_sales = (
            SELECT COUNT(*)
            FROM CustomerSale 
            WHERE offer_id = Offer.id
        )
    """)
    conn.commit()
def update_offer_product():
    # Update ALL rows in Offer_Product:
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE Offer_Product
        SET total_product_purchase_price = quantity_purchased * purchase_price_per_unit
    """)
    print("updated offer product table with total_product_purchase_price")
    conn.commit()
    conn.close()


def update_customer_sale_product():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE CustomerSale_Product
        SET total_product_sale_price = quantity_sold * sell_price_per_unit
    """)
    conn.commit()
    conn.close()



def update_offer_by_offer_id(offer_id):
    """
    Update the Offer row with id=offer_id.
    It recalculates:
      - expected_flip_date: MAX(sell_date) from CustomerSale for that offer.
      - total_sale_price: SUM(total_sale_price) from CustomerSale.
      - sale_complete: 1 if there are no CustomerSale rows with sale_complete=0.
      - investment: SUM(quantity_purchased * purchase_price_per_unit) from Offer_Product.
      - profit: (total_sale_price - investment) if sale_complete is true; otherwise 0.
      - number_of_sales: COUNT(*) of CustomerSale rows for that offer.
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE Offer
        SET 
            expected_flip_date = (
                SELECT MAX(sell_date)
                FROM CustomerSale 
                WHERE offer_id = ?
            ),
            total_sale_price = (
                SELECT COALESCE(SUM(total_sale_price), 0)
                FROM CustomerSale 
                WHERE offer_id = ?
            ),
            sale_complete = CASE 
                                WHEN (SELECT COUNT(*) 
                                      FROM CustomerSale 
                                      WHERE offer_id = ? AND sale_complete = 0) = 0 
                                THEN 1 
                                ELSE 0 
                            END,
            investment = (
                SELECT COALESCE(SUM(quantity_purchased * purchase_price_per_unit), 0)
                FROM Offer_Product 
                WHERE offer_id = ?
            ),
            profit = 
                       
                    (SELECT COALESCE(SUM(total_sale_price), 0)
                     FROM CustomerSale 
                     WHERE offer_id = ?)
                    -
                    (SELECT COALESCE(SUM(quantity_purchased * purchase_price_per_unit), 0)
                     FROM Offer_Product 
                     WHERE offer_id = ?),



            number_of_sales = (
                SELECT COUNT(*)
                FROM CustomerSale 
                WHERE offer_id = ?
            )
        WHERE id = ?
    """, (offer_id, offer_id, offer_id, offer_id, offer_id, offer_id, offer_id, offer_id))
    conn.commit()
    conn.close()
    print("Updated Offer with ID:", offer_id)


def update_offer_product_by_offer_id(offer_id):
    """
    Update all Offer_Product rows for the given offer_id.
    It sets total_product_purchase_price as:
       quantity_purchased * purchase_price_per_unit.
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE Offer_Product
        SET total_product_purchase_price = quantity_purchased * purchase_price_per_unit
        WHERE offer_id = ?
    """, (offer_id,))
    conn.commit()
    conn.close()
    print("updated offer product table with total_product_purchase_price for offer id", offer_id)


def update_customer_sale_product_by_sale_id(sale_id):
    """
    Update the CustomerSale_Product rows for a specific sale.
    It sets total_sale_price as:
       quantity_sold * sell_price_per_unit.
    Here we assume that sale_id is the id of the CustomerSale row
    and that the CustomerSale_Product rows reference that sale via customer_sale_id.
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE CustomerSale_Product
        SET total_product_sale_price = quantity_sold * sell_price_per_unit
        WHERE customer_sale_id = ?
    """, (sale_id,))
    conn.commit()
    conn.close()


def update_customer_sale_by_sale_id(sale_id):
    """
    Update a single CustomerSale row with id=sale_id.
    Mark sale_complete as 1 if its sell_date is today or earlier.
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    today = stdlib_datetime.date.today()
    cursor.execute("""
        UPDATE CustomerSale 
        SET sale_complete = 1 
        WHERE id = ? AND sell_date <= ?
    """, (sale_id, today))
    conn.commit()
    conn.close()


def update_inventory_by_product_id(product_id):
    # Calculate total purchased in grams from Offer_Product.
    purchased_query = """
        SELECT COALESCE(SUM(op.quantity_purchased * u.conversion_factor), 0)
        FROM Offer_Product op
        JOIN Unit u ON op.unit_id = u.unit_id
        WHERE op.product_id = ?
    """
    total_purchased = execute_query(purchased_query, (product_id,), fetch=True)[0][0]
    print(f"Total Purchased for Product {product_id}: {total_purchased}")
    # Calculate total sold in grams from CustomerSale_Product.
    sold_query = """
        SELECT COALESCE(SUM(csp.quantity_sold * u.conversion_factor), 0)
        FROM CustomerSale_Product csp
        JOIN Unit u ON csp.unit_id = u.unit_id
        WHERE csp.product_id = ?
    """
    total_sold = execute_query(sold_query, (product_id,), fetch=True)[0][0]
    print(f"Total Sold for Product {product_id}: {total_sold}")
    # Remaining inventory is purchased minus sold.
    in_inventory = total_purchased - total_sold

    #put this in a sql query for a report
    # print(f"Inventory for Product {product_id}: {in_inventory}")
    # # Calculate total purchase cost from Offer_Product.
    # purchase_cost_query = """
    #     SELECT COALESCE(SUM(op.quantity_purchased * op.purchase_price_per_unit), 0)
    #     FROM Offer_Product op
    #     WHERE op.product_id = ?
    # """
    # total_purchase_cost = execute_query(purchase_cost_query, (product_id,), fetch=True)[0][0]
    # print('total_purchase_cost: ', total_purchase_cost)
    # # Calculate total sale revenue from CustomerSale_Product.
    # sale_revenue_query = """
    #     SELECT COALESCE(SUM(csp.quantity_sold * csp.sell_price_per_unit), 0)
    #     FROM CustomerSale_Product csp
    #     WHERE csp.product_id = ?
    # """
    # total_sale_revenue = execute_query(sale_revenue_query, (product_id,), fetch=True)[0][0]
    # print('total sale revenue: ', total_sale_revenue)
    # # Overall profit is total sale revenue minus total purchase cost.
    # profit_all_time = total_sale_revenue - total_purchase_cost
    #
    # # Profit current could be computed based on current inventory's cost and an average flip price.
    # # Here we use a simple model: if there is any inventory remaining, assume average cost for the remaining inventory.
    # profit_current = profit_all_time  # Adjust as needed if you wish to calculate differently.
    #
    # # Average flip amount (per gram) based on all sold items.
    # avg_flip_query = """
    #     SELECT AVG(csp.sell_price_per_unit)
    #     FROM CustomerSale_Product csp
    #     JOIN Unit u ON csp.unit_id = u.unit_id
    #     WHERE csp.product_id = ?
    # """
    # avg_flip = execute_query(avg_flip_query, (product_id,), fetch=True)[0][0]
    # if avg_flip is None:
    #     avg_flip = 0
    #
    # # Best purchase amount per gram from Offer_Product (lowest purchase_price_per_unit, converted to per gram)
    # best_purchase_query = """
    #     SELECT op.purchase_price_per_unit / u.conversion_factor
    #     FROM Offer_Product op
    #     JOIN Unit u ON op.unit_id = u.unit_id
    #     WHERE op.product_id = ?
    #     ORDER BY op.purchase_price_per_unit / u.conversion_factor ASC
    #     LIMIT 1
    # """
    # best_purchase = execute_query(best_purchase_query, (product_id,), fetch=True)
    # best_purchase_amount = best_purchase[0][0] if best_purchase else 0
    #
    # # For best_purchase_date and best_purchase_contact, assume we want the offer_end_date and contact from the offer
    # best_purchase_info_query = """
    #     SELECT o.offer_end_date, c.name
    #     FROM Offer o
    #     JOIN Offer_Product op ON o.id = op.offer_id
    #     JOIN Contact c ON o.contact_id = c.id
    #     WHERE op.product_id = ?
    #     ORDER BY op.purchase_price_per_unit / (SELECT u.conversion_factor FROM Unit u WHERE op.unit_id = u.unit_id) ASC
    #     LIMIT 1
    # """
    # best_purchase_info = execute_query(best_purchase_info_query, (product_id,), fetch=True)
    # if best_purchase_info:
    #     best_purchase_date, best_purchase_contact = best_purchase_info[0]
    # else:
    #     best_purchase_date, best_purchase_contact = None, None
    #
    # # Best sale amount per gram (highest sell_price_per_unit per conversion)
    # best_sale_query = """
    #     SELECT csp.sell_price_per_unit / u.conversion_factor
    #     FROM CustomerSale_Product csp
    #     JOIN Unit u ON csp.unit_id = u.unit_id
    #     WHERE csp.product_id = ?
    #     ORDER BY csp.sell_price_per_unit / u.conversion_factor DESC
    #     LIMIT 1
    # """
    # best_sale = execute_query(best_sale_query, (product_id,), fetch=True)
    # best_sale_amount = best_sale[0][0] if best_sale else 0
    #
    # # Best sale date corresponding to the best sale amount.
    # best_sale_date_query = """
    #     SELECT cs.sell_date
    #     FROM CustomerSale cs
    #     JOIN CustomerSale_Product csp ON cs.id = csp.customer_sale_id
    #     JOIN Unit u ON csp.unit_id = u.unit_id
    #     WHERE csp.product_id = ?
    #     ORDER BY csp.sell_price_per_unit / u.conversion_factor DESC
    #     LIMIT 1
    # """
    # best_sale_date_result = execute_query(best_sale_date_query, (product_id,), fetch=True)
    # best_sale_date = best_sale_date_result[0][0] if best_sale_date_result else None
    #
    # # Best customer: customer (contact id) who purchased the most (in grams)
    # best_customer_query = """
    #     SELECT c.name, SUM(csp.quantity_sold * u.conversion_factor) as total_qty
    #     FROM CustomerSale_Product csp
    #     JOIN CustomerSale cs ON cs.id = csp.customer_sale_id
    #     JOIN Contact c on cs.contact_id = c.id
    #     JOIN Unit u ON csp.unit_id = u.unit_id
    #     WHERE csp.product_id = ?
    #     GROUP BY cs.contact_id
    #     ORDER BY total_qty DESC
    #     LIMIT 1
    # """
    # best_customer_data = execute_query(best_customer_query, (product_id,), fetch=True)
    # if best_customer_data:
    #     best_customer_id, best_customer_total = best_customer_data[0]
    # else:
    #     best_customer_id, best_customer_total = None, 0

    # Insert or update the Inventory table for this product.
    # We'll use INSERT OR REPLACE so that if an entry exists, it is replaced.
    # query = """
    # INSERT OR REPLACE INTO Inventory (
    #     product_id,
    #     product_amount_purchased_all_time,
    #     product_amount_in_inventory,
    #     product_inventory_purchase_price,
    #     product_amount_sold_all_time,
    #     product_total_sell_amount,
    #     profit_all_time,
    #     profit_current,
    #     avg_flip_amount,
    #     best_purchase_amount,
    #     best_purchase_date,
    #     best_purchase_contact,
    #     best_customer_id,
    #     best_customer_total_purchased,
    #     best_sale_amount,
    #     best_sale_date
    # ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    # """
    # params = (product_id,
    #           total_purchased,
    #           in_inventory,
    #           total_purchase_cost,
    #           total_sold,
    #           total_sale_revenue,
    #           profit_all_time,
    #           profit_current,
    #           avg_flip,
    #           best_purchase_amount,
    #           best_purchase_date,
    #           best_purchase_contact,
    #           best_customer_id,
    #           best_customer_total,
    #           best_sale_amount,
    #           best_sale_date)
    # execute_query(query, params, fetch=False)
    query = """
        INSERT OR REPLACE INTO Inventory (
            product_id,
            product_amount_in_inventory
        ) VALUES (?, ?)
    """
    params = (product_id, in_inventory)

    execute_query(query, params, fetch=False)


def get_products():
    return {str(p[0]): p[1] for p in execute_query(
        "SELECT id, name FROM Product", fetch=True)}

def update_inventory_for_all_products():
    product_dict = get_products()  # returns a dictionary: { "1": "Product A", "2": "Product B", ... }
    print(product_dict)
    for pid in product_dict.keys():
        update_inventory_by_product_id(pid)

def execute_query(query, params=(), fetch=False):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(query, params)
    data = cursor.fetchall() if fetch else None
    conn.commit()
    conn.close()
    return data

def update_database():
    update_customer_sale_product()
    update_customer_sale()
    update_offer_product()
    update_offer()
    update_inventory_for_all_products()




