import logging
import time
import random
import re
import os
import sqlite3
import requests
import sys
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ----------------------------------------------------------------------------
# 1) LOGGING SETUP
# ----------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger()

# ----------------------------------------------------------------------------
# 2) PROXY SETUP (for requests-based fetching)
# ----------------------------------------------------------------------------
proxies_list = [
    "http://123.456.78.90:8080",
    "http://98.76.54.32:3128",
    "http://45.67.89.123:8000",
    "http://209.85.231.53:80",
]

def fetch_random_proxy():
    return random.choice(proxies_list)

def fetch_page_with_requests(url):
    """Fetch a page's HTML using requests and a random proxy; return a BeautifulSoup object."""
    proxy = fetch_random_proxy()
    logger.info(f"[requests] Using proxy {proxy} to fetch {url}")
    try:
        proxies = {"http": proxy, "https": proxy}
        resp = requests.get(url, proxies=proxies, timeout=10)
        resp.raise_for_status()
        return BeautifulSoup(resp.text, "lxml")
    except requests.exceptions.RequestException as e:
        logger.error(f"[requests] Error fetching {url} with proxy {proxy}: {e}")
        return None

# ----------------------------------------------------------------------------
# 3) GET INR → USD CONVERSION RATE
# ----------------------------------------------------------------------------
    #TODO: use units table for conversion
def get_inr_to_usd_rate():
    """Fetch the current INR→USD exchange rate from a free API or return a fallback rate."""
    url = "https://api.exchangerate-api.com/v4/latest/INR"
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        data = r.json()
        rate = data["rates"]["USD"]
        logger.info(f"INR→USD conversion rate: {rate}")
        return rate
    except Exception as e:
        logger.error(f"Error fetching conversion rate: {e}")
        # fallback approximate rate if API fails
        return 0.012

# ----------------------------------------------------------------------------
# 4) SETUP CHROME INCOGNITO
# ----------------------------------------------------------------------------
def setup_chrome_incognito():
    chrome_options = Options()
    chrome_options.add_argument("--incognito")
    # If you want headless, set True:
    headless_mode = False
    if headless_mode:
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    logger.info("Chrome driver launched in incognito mode.")
    return driver

driver = setup_chrome_incognito()


# ----------------------------------------------------------------------------
# 6) SETUP BUSINESS TRACKER DB (MiddleManBusiness/business_tracker7.db)
# ----------------------------------------------------------------------------
if getattr(sys, 'frozen', False):
    # If the application is run as a bundled executable, get the directory of the executable.
    base_dir = os.path.dirname(sys.executable)
else:
    # If it's run as a script, use the directory of the script.
    base_dir = os.path.dirname(__file__)
print(base_dir)
db_file = os.path.join(base_dir, "business_tracker.db")
conn = sqlite3.connect(db_file)
business_cursor = conn.cursor()



# ----------------------------------------------------------------------------
# 7) HELPER FUNCTIONS
# ----------------------------------------------------------------------------
def delay_request():
    time.sleep(random.uniform(3, 6))

def find_desc_sku_value(table_soup, row_label: str) -> str:
    """
    From <table id="desc_sku_tbl">, find <tr> with <td> that matches row_label (case-insensitive).
    Return the next <td> text or '' if not found.
    """
    if not table_soup:
        return ""
    rows = table_soup.find_all("tr", id="desc_sku")
    for row in rows:
        cells = row.find_all("td")
        if len(cells) >= 2:
            label = cells[0].get_text(strip=True).lower()
            if row_label.lower() in label:
                return cells[1].get_text(strip=True)
    return ""

def parse_packaging_size(text: str) -> (float, str):
    """
    If text is something like '10*10 Tablets' or '10 x 6 x 10 Capsules',
    multiply all numeric parts. Return (total_quantity, packaging_unit).
    """
    if not text:
        return 1.0, ""
    normalized = text.replace("*", "x").replace("×", "x")
    nums = re.findall(r'(\d+(?:\.\d+)?)', normalized)
    total_qty = 1.0
    for n in nums:
        total_qty *= float(n)
    # Attempt to find trailing alpha text (e.g. 'Tablets' or 'Capsules')
    m = re.search(r'[A-Za-z]+$', normalized.replace(" ", ""))  # ignoring spaces
    packaging_unit = m.group(0) if m else ""
    return total_qty, packaging_unit



# ----------------------------------------------------------------------------
# 8) SCRAPING LOGIC
# ----------------------------------------------------------------------------
def search_products(search_term: str, inr_to_usd_rate: float):
    """
    1) Go to search page, input search_term, click search.
    2) For each product in the results, open its page, parse details.
    3) Skip if rating < 3.5.
    4) Convert price from INR→USD, store in DBs.
    """
    # 1) Navigate to the search page
    search_page_url = "https://dir.indiamart.com/search.mp?"
    logger.info(f"Navigating to search page: {search_page_url}")
    driver.get(search_page_url)
    time.sleep(3)

    # 2) Perform the search
    try:
        search_input = driver.find_element(By.ID, "search_string1")
        search_input.clear()
        search_input.send_keys(search_term)
        submit_button = driver.find_element(By.ID, "btnSearch")
        submit_button.click()
        logger.info(f"Submitted search for '{search_term}'.")
        time.sleep(5)
    except Exception as e:
        logger.error(f"Error performing search: {e}")
        return

    # 3) Parse the search results
    results_soup = BeautifulSoup(driver.page_source, "lxml")
    product_cards = results_soup.find_all("div", class_="catg_card_txt")
    logger.info(f"Found {len(product_cards)} product cards for search term '{search_term}'")

    for card in product_cards:
        try:
            # (A) Product URL from anchor
            prod_anchor = card.find("a")
            if not prod_anchor:
                logger.warning("No product anchor found in card; skipping.")
                continue
            relative_url = prod_anchor.get("href", "")
            if relative_url.startswith("http"):
                product_url = relative_url
            else:
                product_url = "https://export.indiamart.com/" + relative_url.lstrip("/")

            # (B) Product name from <p id="prdname_x">
            name_tag = card.find("p", id=re.compile(r"^prdname_"))
            product_name = name_tag.get_text(strip=True) if name_tag else "Unknown Product Name"

            logger.info(f"Scraping product: {product_name} => {product_url}")

            # 4) Open product page
            driver.get(product_url)
            delay_request()
            product_soup = BeautifulSoup(driver.page_source, "lxml")

            # 4a) Final product name from <h1 id="prd_name">
            h1_el = product_soup.find("h1", id="prd_name")
            final_product_name = h1_el.get_text(strip=True) if h1_el else product_name

            # 4b) Price from <span id="prc_id"> e.g. "₹ 1000 / Box"
            price_span = product_soup.find("span", id="prc_id")
            final_price_inr = 0.0
            if price_span and "₹" in price_span.get_text():
                raw_price = price_span.get_text(strip=True)
                try:
                    # e.g. "₹ 1000 / Box"
                    raw_no_symbol = raw_price.split("₹", 1)[1].strip()  # "1000 / Box"
                    main_part = raw_no_symbol.split("/")[0].strip()     # "1000"
                    final_price_inr = float(main_part.replace(",", ""))

                except:
                    final_price_inr = 0.0

            # 4c) Packaging size from <table id="desc_sku_tbl">
            desc_table = product_soup.find("table", id="desc_sku_tbl")
            packaging_size_val = find_desc_sku_value(desc_table, "packaging size")

            # parse packaging size => total quantity
            final_quantity, packaging_unit = parse_packaging_size(packaging_size_val)
            if final_quantity <= 0:
                final_quantity = 1.0


            # convert INR→USD
            final_price_usd = final_price_inr * inr_to_usd_rate
            final_price_per_unit_usd = final_price_usd / final_quantity if final_quantity > 0 else 0.0

            # 4d) Supplier name & rating from .cDetails
            cdetails_div = product_soup.find("div", class_="cDetails")
            supplier_name = "Unknown Supplier"
            supplier_rating = 0.0
            if cdetails_div:
                h2_el = cdetails_div.find("h2")
                if h2_el:
                    supplier_name = h2_el.get_text(strip=True)
                rating_el = cdetails_div.find("span", class_="fw7")
                if rating_el:
                    try:
                        supplier_rating = float(rating_el.get_text(strip=True))
                    except:
                        supplier_rating = 0.0
            # skip if rating < 3.5
            if supplier_rating < 3.5:
                logger.info(f"Skipping product '{final_product_name}' - rating {supplier_rating} < 3.5.")
                continue

            # Insert into Contact
            business_cursor.execute(
                "INSERT INTO Contact (type, name, site, phone) VALUES (?, ?, ?, ?)",
                ("supplier", supplier_name, product_url, None)
            )
            contact_id = business_cursor.lastrowid

            # Insert into Offer (store final_price_usd in total_sale_price)
            business_cursor.execute("""
                INSERT INTO Offer (
                    purchase_complete, contact_id, expected_receive_date, expected_flip_date,
                    total_sale_price, profit
                ) VALUES (?, ?, ?, ?, ?, ?)
            """, (False, contact_id, None, None, final_price_usd, None))
            offer_id = business_cursor.lastrowid

            # Insert into Product
            business_cursor.execute(
                "INSERT INTO Product (name) VALUES (?)",
                (final_product_name,)
            )
            product_id = business_cursor.lastrowid

            # Insert into Offer_Product
            business_cursor.execute("""
                INSERT INTO Offer_Product (
                    product_id, offer_id, quantity_purchased,
                    purchase_price_per_unit
                ) VALUES (?, ?, ?, ?)
            """, (
                product_id,
                offer_id,
                final_quantity,
                final_price_per_unit_usd,
            ))
            #TODO: insert into inventory table
            #TODO: use UpdateDatabaseWithCalculatedFields to fill out parts of database
            conn.commit()

            logger.info(f"Business DB inserts done: Contact={contact_id}, Offer={offer_id}, Product={product_id} for '{final_product_name}'")

        except Exception as exc:
            logger.error(f"Error processing product card: {exc}")

def main():
    #TODO: use units table for conversion
    inr_to_usd_rate = get_inr_to_usd_rate()
    search_term = "wheat"
    logger.info(f"Starting search for: {search_term}")
    search_products(search_term, inr_to_usd_rate)
    logger.info("Scraping complete. Leaving browser open for debugging.")
    # driver.quit()

if __name__ == "__main__":
    main()
