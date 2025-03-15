# High-Level Main Script Workflow Description: (addiontal features not in main script at the bottom)
## Database Setup:
When the project is first executed, the CreateDatabase.py script ensures that the SQLite database (business_tracker.db) is created with all the necessary tables, constraints, and relationships.
## Data Insertion via GUI:
The DataInsertGUI.py file launches a Tkinter application where users can input data. This includes adding contacts, offers, products, and customer sales. The GUI dynamically populates dropdowns based on the current database state and performs input validation and unit conversions.
## Updating Calculated Fields:
After data is inserted, functions in UpdateDatabaseWithCalculatedFields.py are called—either automatically or via user actions—to update calculated fields in the database. This ensures that fields like total sale price, profit, inventory, and others are current.
## Report Generation:
The CreateFullReportWithPivotTables.py script (or similar module) queries the database, constructs comprehensive raw data reports using Pandas, and then writes those reports to an Excel file. It also appends additional analysis sheets (such as inventory analytics) and uses xlwings to automatically create pivot tables for summarization and further analysis.
## Excel Pivot Tables for Analysis:
The pivot table functions dynamically extract headers from the raw data sheet and build pivot tables on separate sheets, allowing users to quickly analyze sales by customer, offer, product, and date. The code is designed to be robust by handling header mismatches, ensuring contiguous data ranges, and preventing conflicts with extra workbooks.

# Detailed Main Script Execution
## 1. Database Creation (CreateDatabase.py)
#### create_database
This file sets up the SQLite database (named “business_tracker.db”) with all the required tables and relationships. The tables created are:
 ##### Contact:
 Stores information about suppliers, customers, and middlemen. It includes fields such as id (an auto‑increment primary key), name, site, phone, and type (with a CHECK constraint to allow only 'supplier', 'middleman', or 'customer'). Unique constraints ensure that the combination of id and name, or id and phone, remains unique.
 ##### Offer:
 Represents an offer made by a supplier. It has fields like purchase_complete and sale_complete (both booleans with default values), dates (offer_start_date, offer_end_date, expected_receive_date, expected_flip_date), and financial metrics such as investment, total_sale_price, and profit. It also stores number_of_sales. The contact_id is a foreign key linking back to the Contact table.
 ##### Offer_Product:
 This is a junction table connecting an offer to one or more products. It stores the quantity purchased and the purchase price per unit. It also holds a calculated field total_product_purchase_price (which is later updated via a separate function) and a unit_id field that references a Unit table. A composite primary key on (offer_id, product_id) prevents duplicate entries.
 ##### CustomerSale:
 This table records a sale to a customer. It contains fields such as sell_date, total_sale_price, and a flag sale_complete that indicates whether the sale has been finalized (for example, if the sell_date is today or earlier). It also stores customer information and links to the Offer and Contact tables via foreign keys.
 ##### CustomerSale_Product:
 Similar in concept to Offer_Product, this junction table links a CustomerSale with one or more products. It records the quantity sold, the selling price per unit, and a calculated field total_product_sale_price. It also references the Unit table via unit_id.
 ##### Inventory:
 A simple table that keeps track of each product’s current inventory level (in grams, for instance). It has a primary key of product_id and a field for the current amount in inventory.
 ##### Unit:
 This table stores unit names (like “g”, “oz”, “kg”, etc.) along with a conversion factor (typically to grams) and a field for a related unit. Although the original design allowed for a foreign key to Product, the current design omits a product_id so that units are maintained globally.
 ##### Product:
 Stores products with fields for an auto‑increment id and a product name. It also includes a foreign key (unit_id) referencing the Unit table so that each product is associated with a default unit.
#### add_agriculture_data
A simple script to insert fake data into the database.

## 2. Data Insertion and GUI (DataInsertGUI.py)
This module implements a Tkinter‑based graphical interface for entering and modifying business data in the system. 
### The main window features 4 seperate tabs:
##### Tab 1 - Offers:
In the Offers tab, the user selects a supplier from a dynamically refreshed dropdown (populated from the Contact table), enters offer dates, and selects one or more products by checking boxes; for each selected product the user must provide a purchase quantity, unit price, and unit (entered as text or selected from a dropdown that references the Unit table). When an offer is submitted, the script inserts the new offer into the Offer table and adds the product details into the Offer_Product table. 
##### Tab 2 - Customer Sales:
A tab where the user selects an existing offer and a customer, then specifies the sale quantities, selling prices, and units for one or more products. Before inserting a sale, the code uses a conversion function (convert_units) to check that the sale quantities—provided in possibly different units—do not exceed the available purchased amounts (without altering the raw sale quantity used to compute the total sale price).
##### Tab 3 - Contacts:
A tab for adding new contacts (suppliers, customers, or middlemen) with appropriate fields.
##### Tab 4 - Edit Offer:
A tab that enables the user to load an existing offer’s details (including its associated products), manually adjust quantities, prices, or units, and then update the records in the database.
### Data Handling 
#### Dynamic Dropdowns and Data Refresh:
 Throughout the interface, various refresh functions update dropdown menus and product lists (both for adding new data and editing existing records) to ensure they reflect the current state of the database. After each data insertion or update, the GUI calls the update_database() function (described below) from the UpdateDatabaseWithCalculatedFields module to recalculate all derived fields and update inventory levels accordingly. Error handling is built in via Tkinter’s messageboxes, and the interface preserves any manual state entered during edits to ensure smooth user interactions.
#### Error Handling and Validation:
The code uses Tkinter’s messagebox to display errors (e.g., if required fields are missing or numeric values are invalid). It also performs unit conversion checks to ensure that sales do not exceed available purchased quantities.
##### Update Calculated Fields (UpdateDatabaseWithCalculatedFields.py - update_database() function)
This module contains a suite of functions that recalculate and update various computed fields throughout the database after new data is inserted or modified. Key functions include:
###### update_customer_sale() and update_customer_sale_product(): 
These functions mark customer sales as complete (if the sale date is today or earlier) and recalculate the total sale price for each sale by multiplying the quantities sold by their selling prices.
###### update_offer() and update_offer_by_offer_id(): 
These functions update an offer’s overall data by computing the expected flip date (the latest sale date among associated customer sales), summing up all sale prices to get the total sale price, counting the number of sales, and determining whether the offer is complete (i.e. if all its customer sales are finalized). They also calculate the investment by summing the product of quantities purchased and purchase prices from the Offer_Product records, and then derive the profit as the difference between total sale price and investment (when all sales are complete).
###### update_offer_product() and update_offer_product_by_offer_id(): 
These functions recalculate the total product purchase price (i.e. quantity_purchased multiplied by purchase_price_per_unit) for each record in the Offer_Product table.
###### update_inventory_by_product_id() and update_inventory_for_all_products(): 
These functions determine the current inventory for each product by converting all purchased and sold quantities to a common base unit (using conversion factors from the Unit table) and then subtracting the total sold from the total purchased.
###### update_customer_sale_product_by_sale_id() and update_customer_sale_by_sale_id(): 
These update functions recalculate sale totals and completion flags for individual sales.
The update_database() function is a convenience wrapper that sequentially calls the above update functions to bring the entire database’s calculated fields (including offer totals, profit, sales completion, and inventory levels) up to date. Note that this update process is not part of the main application’s execution by default—it must be run independently when a full recalculation is needed.

## 4. Full Report and Pivot Tables (CreateFullReportWithPivotTables.py)
This file pulls data from the database, creates a full raw data report, and then uses pivot tables for summarization. Key steps include:
#### calculate_and_update_inventory_stats
This function creates a basic sheet on the excel workbook for current inventory statistics. It does this by executing SQL queries to retrieve data from the Offer, Offer_Product, and CustomerSale (and CustomerSale_Product) tables. The queries join the necessary tables (and sometimes use subqueries) to retrieve computed values (or use the pre‑calculated fields that were updated earlier).
Pandas DataFrames are built from these queries to represent raw data in a cascading (or tabular) format.
#### generate_excel_report
The raw data is exported to an Excel file (full_analytics_report.xlsx) with a designated sheet (usually “Sheet1”).
#### create_pivot_with_xlwings:
The functions create pivot tables using Excel’s COM interface via xlwings. Options such as add_book=False prevent the creation of extra blank workbooks, and the pivot tables are created on newly added or cleared sheets.

## SQL Features Utilized
### Foreign Keys and Cascading Deletes:
Enforced in tables like Offer, Offer_Product, CustomerSale, and CustomerSale_Product to maintain referential integrity.
###Constraints (CHECK and UNIQUE):
Used in the Contact table (to restrict the type field), Offer_Product and CustomerSale_Product (to ensure quantities are positive), and elsewhere.
## Subquerie:
In the report queries and update functions, subqueries are used extensively to calculate sums, averages, and other aggregates (e.g., total_sale_price, investment, profit).
### Aggregate Functions:
Functions like SUM, AVG, MIN, MAX, and COUNT are used to compute financial totals and inventory statistics.
### PRAGMA Settings:
Specifically, enabling foreign key constraints with PRAGMA foreign_keys = ON;.

# Conclusion
When executed, this project creates a complete business tracker system that:
Sets up a structured SQLite database,
Provides a GUI for entering and editing data,
Automatically updates calculated fields,
Generates detailed Excel reports, and
Creates pivot tables for dynamic analysis.
Each module plays a specific role, and together they form an integrated solution for tracking offers, sales, and inventory with robust reporting and analysis capabilities.

# Additional Files for Executing Independently
## Web Scraping (IndiaMartAgriculturalScraper.py)
The scraping module is a standalone script that programmatically gathers product and supplier information from IndiaMART. Using Selenium WebDriver (launched in incognito mode) alongside BeautifulSoup for HTML parsing, the script navigates to IndiaMART’s search page, inputs a user-specified search term, and iterates over the resulting product listings. For each product, it opens the detailed product page to extract key details such as the final product name, price (in INR), packaging size (from which the total available quantity is calculated by multiplying the numeric factors), and supplier details—including the supplier’s name and rating. The module further retrieves real-time INR-to-USD conversion rates via an external API, converting prices accordingly and applying any necessary unit conversion logic (e.g., translating kilograms to milligrams when validating available quantities). Products with a supplier rating below 3.5 are automatically skipped. Finally, the extracted data is inserted into the business tracker database by creating new records in the Contact, Offer, Product, and Offer_Product tables. This process leverages randomized proxy usage and built-in delays to mimic natural browsing behavior and ensure reliable data collection. (Note: At present, this scraping functionality is not integrated into the main application flow and must be executed as a separate process.)
