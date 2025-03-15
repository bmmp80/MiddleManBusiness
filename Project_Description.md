# Main Script Execution
## 1. Database Creation (CreateDatabase.py)
This file sets up the SQLite database (named “business_tracker.db”) with all the required tables and relationships. Key points include:

SQLite Connection & PRAGMA Settings:
The script opens a connection to the database file and enables foreign key constraints via
PRAGMA foreign_keys = ON;
This ensures that relationships between tables are enforced (for example, deleting a Contact will cascade and delete associated Offers).

Table Creation:
Several tables are created using SQL’s CREATE TABLE IF NOT EXISTS syntax. For example:

Contact:
Stores information about suppliers, customers, and middlemen. It includes fields such as id (an auto‑increment primary key), name, site, phone, and type (with a CHECK constraint to allow only 'supplier', 'middleman', or 'customer'). Unique constraints ensure that the combination of id and name, or id and phone, remains unique.

Offer:
Represents an offer made by a supplier. It has fields like purchase_complete and sale_complete (both booleans with default values), dates (offer_start_date, offer_end_date, expected_receive_date, expected_flip_date), and financial metrics such as investment, total_sale_price, and profit. It also stores number_of_sales. The contact_id is a foreign key linking back to the Contact table.

Offer_Product:
This is a junction table connecting an offer to one or more products. It stores the quantity purchased and the purchase price per unit. It also holds a calculated field total_product_purchase_price (which is later updated via a separate function) and a unit_id field that references a Unit table. A composite primary key on (offer_id, product_id) prevents duplicate entries.

CustomerSale:
This table records a sale to a customer. It contains fields such as sell_date, total_sale_price, and a flag sale_complete that indicates whether the sale has been finalized (for example, if the sell_date is today or earlier). It also stores customer information and links to the Offer and Contact tables via foreign keys.

CustomerSale_Product:
Similar in concept to Offer_Product, this junction table links a CustomerSale with one or more products. It records the quantity sold, the selling price per unit, and a calculated field total_product_sale_price. It also references the Unit table via unit_id.

Inventory:
A simple table that keeps track of each product’s current inventory level (in grams, for instance). It has a primary key of product_id and a field for the current amount in inventory.

Unit:
This table stores unit names (like “g”, “oz”, “kg”, etc.) along with a conversion factor (typically to grams) and a field for a related unit. Although the original design allowed for a foreign key to Product, the current design omits a product_id so that units are maintained globally.

Product:
Stores products with fields for an auto‑increment id and a product name. It also includes a foreign key (unit_id) referencing the Unit table so that each product is associated with a default unit.

SQL Features Utilized:
The project uses several SQL features:

Primary Keys and Auto‑Increment: To uniquely identify rows.
Foreign Keys with CASCADE: Enforcing referential integrity between tables.
CHECK Constraints: For example, in the Contact table to restrict the type field.
UNIQUE Constraints: Preventing duplicate data.
Default Values: For boolean fields and financial figures.
## 2. Data Insertion and GUI (DataInsertGUI.py)
This file contains a Tkinter‑based graphical user interface that lets users add data into the system. Key elements include:

GUI Setup with Tkinter and ttk:
The application creates a main window with a notebook widget (tabbed interface). Different tabs allow the user to add:

Offers:
A tab that allows the user to select a supplier (from a dropdown populated from the Contact table), enter offer dates, and select products (with checkboxes, quantity, purchase price, and unit entry fields). When an offer is submitted, it inserts data into the Offer and Offer_Product tables.

Contacts:
A tab for adding new contacts (suppliers, customers, or middlemen) with appropriate fields.

Customer Sales:
A tab where a user selects an offer, a customer, and then selects products (with sale quantities, sell prices, and units). It performs necessary conversions (using the convert_units function) to check that sales do not exceed what was purchased.

Edit Offer:
A tab that allows editing existing offers. It populates the fields and product list based on the selected offer and lets the user update them.

Dynamic Dropdowns and Data Refresh:
Functions like refresh_dropdowns(), refresh_product_list(), and refresh_sale_product_list() update the GUI controls based on the current state of the database. This ensures that when new contacts, offers, or products are added, the dropdowns and checklists reflect the changes.

Unit Selection:
In later versions, the product unit fields are replaced by dropdowns that reference the Unit table (using unit_id). Buttons to add a new unit allow the user to specify unit conversion factors and related units.

Error Handling and Validation:
The code uses Tkinter’s messagebox to display errors (e.g., if required fields are missing or numeric values are invalid). It also performs unit conversion checks to ensure that sales do not exceed available purchased quantities.

## 3. Update Calculated Fields (UpdateDatabaseWithCalculatedFields.py)
This module contains functions that update various calculated fields in the database after data insertion. Examples include:

update_customer_sale() and update_customer_sale_product():
These functions update the sale_complete flag in the CustomerSale table and calculate the total product sale price in CustomerSale_Product by multiplying quantity sold by the sell price per unit.

update_offer() and update_offer_by_offer_id():
These functions update an Offer’s computed fields, such as:

expected_flip_date (the latest sell_date among related CustomerSales),
total_sale_price (the sum of sale prices),
sale_complete (whether all associated sales are complete),
investment (the sum of the product of quantity purchased and purchase price per unit from Offer_Product),
profit (calculated as total sale price minus investment, provided that the sale is complete),
number_of_sales (the count of CustomerSales for the offer).
update_offer_product() and update_customer_sale_product():
These functions update calculated fields in the junction tables (Offer_Product and CustomerSale_Product) such as total_product_purchase_price and total_product_sale_price.

update_inventory_by_product_id():
This function calculates the inventory for a given product by converting all purchased quantities (using the conversion factor from the Unit table) and subtracting the sold quantities. It updates the Inventory table accordingly.

## 4. Full Report and Pivot Tables (CreateFullReportWithPivotTables.py)
This file pulls data from the database, creates a full raw data report, and then uses pivot tables for summarization. Key steps include:

Data Extraction Using SQL and Pandas:
The script executes SQL queries to retrieve data from the Offer, Offer_Product, and CustomerSale (and CustomerSale_Product) tables. The queries join the necessary tables (and sometimes use subqueries) to retrieve computed values (or use the pre‑calculated fields that were updated earlier).
Pandas DataFrames are built from these queries to represent raw data in a cascading (or tabular) format.

## 5 Excel Report Generation:
The raw data is exported to an Excel file (full_analytics_report.xlsx) with a designated sheet (usually “Sheet1”).
Then, the script uses xlwings to create pivot tables on additional sheets. One pivot table might group by customer and date to summarize sales, while another might group by offer and product to summarize quantities sold and revenue.
The pivot table functions dynamically extract the header names from the raw data sheet so that the field names match exactly what’s in the Excel file.
The script also uses ExcelWriter in append mode (with openpyxl) to add new sheets (for example, an “Inventory Analytics” sheet) without overwriting existing data.

xlwings Automation:
The functions create pivot tables using Excel’s COM interface via xlwings. Options such as add_book=False prevent the creation of extra blank workbooks, and the pivot tables are created on newly added or cleared sheets.

# Overall Main Script Workflow
Database Setup:
When the project is first executed, the CreateDatabase.py script ensures that the SQLite database (business_tracker.db) is created with all the necessary tables, constraints, and relationships.

Data Insertion via GUI:
The DataInsertGUI.py file launches a Tkinter application where users can input data. This includes adding contacts, offers, products, and customer sales. The GUI dynamically populates dropdowns based on the current database state and performs input validation and unit conversions.

Updating Calculated Fields:
After data is inserted, functions in UpdateDatabaseWithCalculatedFields.py are called—either automatically or via user actions—to update calculated fields in the database. This ensures that fields like total sale price, profit, inventory, and others are current.

Report Generation:
The CreateFullReportWithPivotTables.py script (or similar module) queries the database, constructs comprehensive raw data reports using Pandas, and then writes those reports to an Excel file. It also appends additional analysis sheets (such as inventory analytics) and uses xlwings to automatically create pivot tables for summarization and further analysis.

Excel Pivot Tables for Analysis:
The pivot table functions dynamically extract headers from the raw data sheet and build pivot tables on separate sheets, allowing users to quickly analyze sales by customer, offer, product, and date. The code is designed to be robust by handling header mismatches, ensuring contiguous data ranges, and preventing conflicts with extra workbooks.

SQL Features Utilized
Foreign Keys and Cascading Deletes:
Enforced in tables like Offer, Offer_Product, CustomerSale, and CustomerSale_Product to maintain referential integrity.

Constraints (CHECK and UNIQUE):
Used in the Contact table (to restrict the type field), Offer_Product and CustomerSale_Product (to ensure quantities are positive), and elsewhere.

Subqueries:
In the report queries and update functions, subqueries are used extensively to calculate sums, averages, and other aggregates (e.g., total_sale_price, investment, profit).

Aggregate Functions:
Functions like SUM, AVG, MIN, MAX, and COUNT are used to compute financial totals and inventory statistics.

PRAGMA Settings:
Specifically, enabling foreign key constraints with PRAGMA foreign_keys = ON;.

Conclusion
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
