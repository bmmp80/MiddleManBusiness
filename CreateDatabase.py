import sqlite3
import os
import sys
print("Using Python:", sys.executable)
print("Python version:", sys.version)
print("sys.path:", sys.path)

import pandas as pd
# Define the correct database directory and path
print('directory name: ', os.path.dirname)
# test
# db_dir = "C:/Users/brand/PycharmProjects/MiddleManBusiness/Attempt7"
# db_file = os.path.join(db_dir, "business_tracker.db")
def create_database():
    # executable

    if getattr(sys, 'frozen', False):
        # If the application is run as a bundled executable, get the directory of the executable.
        base_dir = os.path.dirname(sys.executable)
    else:
        # If it's run as a script, use the directory of the script.
        base_dir = os.path.dirname(__file__)
    print(base_dir)
    db_file = os.path.join(base_dir, "business_tracker.db")

    # Ensure the directory exists
    # os.makedirs(db_dir, exist_ok=True)

    # Connect to the SQLite database
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    # Enable foreign key constraints
    cursor.execute("PRAGMA foreign_keys = ON;")

    # Create Contact table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Contact (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            site TEXT,
            phone TEXT,
            type TEXT NOT NULL CHECK (type IN ('supplier', 'middleman', 'customer')),
            UNIQUE(id, name),
            UNIQUE(id, phone)
        );
    """)

    # Create Offer table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Offer (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            purchase_complete boolean default false,
            sale_complete boolean default false,
            offer_start_date TEXT,
            offer_end_date TEXT,
            expected_receive_date TEXT,
            expected_flip_date TEXT,
            investment REAL DEFAULT NULL,
            total_sale_price REAL DEFAULT 0,
            profit REAL DEFAULT 0,
            contact_id INTEGER NOT NULL,
            number_of_sales INTEGER default 0,
            FOREIGN KEY (contact_id) REFERENCES Contact (id) ON DELETE CASCADE           
        );
    """)

    # Create Product table with only the product name
    # cursor.execute("""
    #     CREATE TABLE IF NOT EXISTS Product (
    #         id INTEGER PRIMARY KEY AUTOINCREMENT,
    #         name TEXT NOT NULL
    #     );
    # """)

    # Create Offer_Product table with purchase_price_per_unit
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Offer_Product (
            offer_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL,
            quantity_purchased INTEGER NOT NULL CHECK(quantity_purchased > 0),
            purchase_price_per_unit REAL NOT NULL,
            total_product_purchase_price REAL,
            unit_id REFERENCES Unit (unit_id) ON DELETE CASCADE,
            PRIMARY KEY (offer_id, product_id),
            FOREIGN KEY (offer_id) REFERENCES Offer (id) ON DELETE CASCADE,
            FOREIGN KEY (product_id) REFERENCES Product (id) ON DELETE CASCADE
        );
    """)

    # Create CustomerSale table (unchanged from before)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS CustomerSale (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sale_complete boolean default false,
            sell_date TEXT,
            offer_id INTEGER NOT NULL,
            contact_id INTEGER NOT NULL,
            quantity_sold INTEGER NOT NULL CHECK(quantity_sold > 0),
            total_sale_price REAL DEFAULT 0,
            customer TEXT NOT NULL,
            customer_phone TEXT,
            FOREIGN KEY(customer) REFERENCES Contact (name),
            FOREIGN KEY (offer_id) REFERENCES Offer (id) ON DELETE CASCADE,
            FOREIGN KEY (contact_id, customer) REFERENCES Contact (id, name),
            FOREIGN KEY (contact_id, customer_phone) REFERENCES Contact (id, phone)
        );
    """)

    # Create CustomerSale_Product table with sell_price_per_unit
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS CustomerSale_Product (
            customer_sale_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL,
            quantity_sold INTEGER NOT NULL CHECK(quantity_sold > 0),
            sell_price_per_unit REAL NOT NULL,
            unit_id REFERENCES Unit (unit_id) ON DELETE CASCADE,
            total_product_sale_price REAL DEFAULT 0,
            PRIMARY KEY (customer_sale_id, product_id),
            FOREIGN KEY (customer_sale_id) REFERENCES CustomerSale (id) ON DELETE CASCADE,
            FOREIGN KEY (product_id) REFERENCES Product (id) ON DELETE CASCADE
        );
    """)

    # cursor.execute("""
    # CREATE TABLE IF NOT EXISTS Unit (
    #     unit_id INTEGER PRIMARY KEY AUTOINCREMENT,
    #     unit_name TEXT NOT NULL,
    #     conversion_factor REAL NOT NULL,  -- factor to convert this unit to grams
    #     related_unit_id INTEGER,
    #     FOREIGN KEY(product_id) REFERENCES Product(id) ON DELETE CASCADE
    # );
    # """)
    # conn.commit()

    # For product_id = 1, insert units: "g", "oz", "lb", "kg"
    # We'll assume:
    #   1 g = 1 g (conversion_factor = 1)
    #   1 oz = 28 g (conversion_factor = 28)
    #   1 lb = 453.592 g (conversion_factor = 453.592)
    #   1 kg = 1000 g (conversion_factor = 1000)
    #
    # We'll also set related_unit_id for "g" <--> "oz" and "lb" <--> "kg".
    # Since the related_unit_id is a foreign key, you must insert one row,
    # then use its id for the other. For simplicity, we'll insert them in order
    # and then update the related_unit_id afterward.

    product_id = 1  # For example

    # Insert "g" and "oz"
    # cursor.execute("INSERT INTO Unit (product_id, unit_name, conversion_factor) VALUES (?, ?, ?)", (product_id, "g", 1))
    # g_id = cursor.lastrowid
    # cursor.execute("INSERT INTO Unit (product_id, unit_name, conversion_factor) VALUES (?, ?, ?)",
    #                (product_id, "oz", 28))
    # oz_id = cursor.lastrowid

    # Link them as related units:
    # cursor.execute("UPDATE Unit SET related_unit_id = ? WHERE unit_id = ?", (oz_id, g_id))
    # cursor.execute("UPDATE Unit SET related_unit_id = ? WHERE unit_id = ?", (g_id, oz_id))

    # Insert "lb" and "kg"
    # cursor.execute("INSERT INTO Unit (product_id, unit_name, conversion_factor) VALUES (?, ?, ?)",
    #                (product_id, "lb", 453.592))
    # lb_id = cursor.lastrowid
    # cursor.execute("INSERT INTO Unit (product_id, unit_name, conversion_factor) VALUES (?, ?, ?)",
    #                (product_id, "kg", 1000))
    # kg_id = cursor.lastrowid

    # Link them as related units:
    # cursor.execute("UPDATE Unit SET related_unit_id = ? WHERE unit_id = ?", (kg_id, lb_id))
    # cursor.execute("UPDATE Unit SET related_unit_id = ? WHERE unit_id = ?", (lb_id, kg_id))
    conn.commit()
    cursor.execute("""
            CREATE TABLE IF NOT EXISTS Inventory (
                product_id INTEGER PRIMARY KEY,
                product_amount_in_inventory REAL DEFAULT 0,
                FOREIGN KEY(product_id) REFERENCES Product(id) ON DELETE CASCADE
            );
        """)

    # cursor.execute("CREATE TEMP TABLE temp_Inventory AS SELECT product_id, product_amount_in_inventory FROM Inventory;")
    # cursor.execute(
    #     "CREATE TEMP TABLE temp_Unit AS SELECT unit_id, unit_name, conversion_factor, related_unit_id FROM Unit;")
    # cursor.execute("CREATE TEMP TABLE temp_Product AS SELECT id, name, unit_id FROM Product;")
    #
    # ### Step 2: Drop the Original Tables ###
    # cursor.execute("DROP TABLE IF EXISTS Inventory;")
    # cursor.execute("DROP TABLE IF EXISTS Unit;")
    # cursor.execute("DROP TABLE IF EXISTS Product;")







    ### Step 4: Restore Data in Correct Order ###

    # Restore Unit first (since Product depends on it)
    cursor.execute("""
            CREATE TABLE IF NOT EXISTS Unit (
                unit_id INTEGER PRIMARY KEY AUTOINCREMENT,
                unit_name TEXT NOT NULL,
                conversion_factor REAL NOT NULL,
                related_unit_id INTEGER
            );
        """)
    cursor.execute("""
            CREATE TABLE IF NOT EXISTS Product (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                unit_id INTEGER,
                FOREIGN KEY(unit_id) REFERENCES Unit(unit_id) ON DELETE CASCADE
            );
        """)

    # cursor.execute("""
    # CREATE TABLE IF NOT EXISTS Inventory (
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
    #     FOREIGN KEY(product_id) REFERENCES Product(id) ON DELETE CASCADE
    # );
    # """)
    conn.commit()

def add_algriculture_data():
    import sqlite3

    if getattr(sys, 'frozen', False):
        # If the application is run as a bundled executable, get the directory of the executable.
        base_dir = os.path.dirname(sys.executable)
    else:
        # If it's run as a script, use the directory of the script.
        base_dir = os.path.dirname(__file__)
    print(base_dir)
    db_file = os.path.join(base_dir, "business_tracker.db")

    # Connect to the SQLite database
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    # Insert static data into Contact table (no dynamically calculated fields)
    contact_data = [
        (1, "Green Farms", "www.greenfarms.com", "123-456-7890", "supplier"),
        (2, "John's Market", "www.johnsmarket.com", "987-654-3210", "customer"),
        (3, "AgriMiddle Ltd.", "www.agrimiddle.com", "555-666-7777", "middleman"),
        (4, "Fresh Groceries", "www.freshgroceries.com", "222-333-4444", "customer"),
        (5, "Organic Supplies", "www.organicsupplies.com", "888-999-0000", "supplier")
    ]
    cursor.executemany("INSERT INTO Contact (id, name, site, phone, type) VALUES (?, ?, ?, ?, ?)", contact_data)

    # Insert static data into Product table
    product_data = [
        (1, "Wheat", 3),
        (2, "Corn", 3),
        (3, "Tomatoes", 3),
        (4, "Carrots", 3),
        (5, "Apples", 3)
    ]
    cursor.executemany("INSERT INTO Product (id, name, unit_id) VALUES (?, ?, ?)", product_data)

    # Insert static data into Unit table
    unit_data = [
        (1, "g", 1, None),
        (2, "oz", 28, 1),
        (3, "kg", 1000, 1),
        (4, "lb", 16, 2),
        (5, "ton", 2000, 4),
        (6, "litre", 1, None),
        (7, "ml", 1000, 6)
    ]
    cursor.executemany("INSERT INTO Unit (unit_id, unit_name, conversion_factor, related_unit_id) VALUES (?, ?, ?, ?)",
                       unit_data)

    # # Insert inventory
    # inventory_data = [
    #     (1, 1000.0),
    #     (2, 800.0),
    #     (3, 600.0),
    #     (4, 500.0),
    #     (5, 750.0)
    # ]
    # cursor.executemany("INSERT INTO Inventory (product_id, product_amount_in_inventory) VALUES (?, ?)", inventory_data)

    # Insert Offer table **excluding fields that are updated dynamically**
    offer_data = [
        (1, 1, "2025-02-01", "2025-02-10", "2025-02-01", "2025-02-15"),
        (2, 1, "2025-02-05", "2025-02-12", "2025-02-05", "2025-02-18"),
        (3, 5, "2025-02-08", "2025-02-14", "2025-02-08", "2025-02-20")
    ]
    cursor.executemany("""
        INSERT INTO Offer (id, contact_id, expected_flip_date, expected_receive_date, offer_start_date, offer_end_date)
        VALUES (?, ?, ?, ?, ?, ?)
    """, offer_data)

    # Insert Offer_Product table **excluding total_product_purchase_price (calculated separately)**
    offer_product_data = [
        (1, 1, 500, 10.0, 3),
        (1, 2, 300, 15.0, 3),
        (2, 3, 400, 8.0, 3),
        (3, 4, 600, 12.0, 3),
        (3, 5, 500, 11.0, 3)
    ]
    cursor.executemany("""
        INSERT INTO Offer_Product (offer_id, product_id, quantity_purchased, purchase_price_per_unit, unit_id)
        VALUES (?, ?, ?, ?, ?)
    """, offer_product_data)

    # Insert CustomerSale table **excluding sale_complete (calculated dynamically)**
    customer_sale_data = [
        (1, 2, "John's Market", "987-654-3210", "2025-02-20", 1, 200),
        (2, 2, "John's Market", "987-654-3210", "2025-02-21", 1, 150),
        (3, 4, "Fresh Groceries", "222-333-4444", "2025-02-22", 2, 250),
        (4, 4, "Fresh Groceries", "222-333-4444", "2025-02-23", 2, 100),
        (5, 2, "John's Market", "987-654-3210", "2025-02-24", 3, 300),
        (6, 4, "Fresh Groceries", "222-333-4444", "2025-02-25", 3, 350)
    ]
    cursor.executemany("""
        INSERT INTO CustomerSale (id, contact_id, customer, customer_phone, sell_date, offer_id, quantity_sold)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, customer_sale_data)

    # Insert CustomerSale_Product table **excluding total_product_sale_price (calculated separately)**
    customer_sale_product_data = [
        (1, 1, 200, 20.0, 3),
        (2, 1, 150, 20.0, 3),
        (3, 3, 250, 20.0, 3),
        (4, 3, 100, 20.0, 3),
        (5, 4, 300, 20.0, 3),
        (6, 5, 350, 20.0, 3)
    ]
    cursor.executemany("""
        INSERT INTO CustomerSale_Product (customer_sale_id, product_id, quantity_sold, sell_price_per_unit, unit_id)
        VALUES (?, ?, ?, ?, ?)
    """, customer_sale_product_data)

    # Commit and close
    conn.commit()
    conn.close()

    print("Static data inserted. Dynamic fields will be updated separately.")

    # Display the expanded dataset script


create_database()
add_algriculture_data()