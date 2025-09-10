from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import re
import time
from datetime import datetime

# ========================
# CONFIGURATION
# ========================
CATEGORIES = [
    ("Cement", "https://www.randtech.co.ke/product-category/flooring/cement/"),
    ("Paint", "https://www.randtech.co.ke/product-category/paint/"),
    ("Solar Lights", "https://www.randtech.co.ke/product-category/electricals/solar-lights/"),
    ("Plumbing", "https://www.randtech.co.ke/product-category/plumbing/"),
    ("Tanks", "https://www.randtech.co.ke/product-category/building-materials/tanks/")
]

# Format: day-month-year
TIMESTAMP = datetime.now().strftime('%d-%m-%Y')
FILE_NAME = f"randtech_{TIMESTAMP}.xlsx"

# ========================
# SELENIUM DRIVER SETUP
# ========================
def setup_driver():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(180)
    return driver

# ========================
# SCRAPING FUNCTION
# ========================
def scrape_category(driver, category_name, category_url):
    print(f"\n🔍 Scraping category: {category_name}")
    products = []
    page_number = 1

    while True:
        page_url = category_url if page_number == 1 else f"{category_url}page/{page_number}/"
        try:
            driver.get(page_url)
        except Exception as e:
            print(f"⚠️ Error loading page {page_number}: {str(e)} — skipping.")
            break

        try:
            # wait for product elements to load
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.product-small"))
            )
        except:
            print(f"✅ No more products found on page {page_number}.")
            break

        product_cards = driver.find_elements(By.CSS_SELECTOR, "div.product-small")

        if not product_cards:
            print(f"✅ No more products found on page {page_number}.")
            break

        print(f"✅ Found {len(product_cards)} products on page {page_number}")

        for card in product_cards:
            try:
                # product name
                try:
                    name = card.find_element(By.CSS_SELECTOR, ".box-text .name.product-title").text
                except:
                    name = "N/A"

                # product price
                try:
                    price = card.find_element(By.CSS_SELECTOR, "span.woocommerce-Price-amount").text
                except:
                    price = "N/A"

                # product link
                try:
                    link = card.find_element(By.CSS_SELECTOR, "a").get_attribute("href")
                except:
                    link = "N/A"

                # product image
                try:
                    image = card.find_element(By.CSS_SELECTOR, "img").get_attribute("src")
                except:
                    image = "N/A"

                products.append({
                    "Category": category_name,
                    "Product Name": name,
                    "Price": price,
                    "Link": link,
                    "Image": image
                })

            except Exception as e:
                print(f"⚠️ Error extracting product: {str(e)}")
                continue

        page_number += 1
        time.sleep(2)  # small delay to avoid blocking

    return products

# ========================
# SPLIT PRODUCT DETAILS
# ========================
def split_product_details(full_name):
    if not isinstance(full_name, str):
        return pd.Series({'Product_Name': None, 'Quantity': None})

    quantity_pattern = r'(?i)(\d+(?:\.\d+)?)\s*(ml|l|g|kg|pcs|pack|pieces|grams|kilos|ltr|litres)'
    quantity_match = re.search(quantity_pattern, full_name)

    if quantity_match:
        qty_value = quantity_match.group(1)
        qty_unit = quantity_match.group(2).upper()
        quantity = f"{qty_value}{qty_unit}"
        product_name = re.split(quantity_pattern, full_name, maxsplit=1, flags=re.IGNORECASE)[0].strip()
    else:
        product_name = full_name
        quantity = None

    return pd.Series({
        'Product_Name': product_name,
        'Quantity': quantity
    })

# ========================
# MAIN FUNCTION
# ========================
def main():
    driver = setup_driver()
    all_products = []

    try:
        for cat_name, cat_url in CATEGORIES:
            products = scrape_category(driver, cat_name, cat_url)
            all_products.extend(products)

        if all_products:
            df = pd.DataFrame(all_products)

            # Split product name & quantity
            details_df = df['Product Name'].apply(split_product_details)
            df_final = pd.concat([df[['Category']], details_df, df[['Price', 'Link', 'Image']]], axis=1)

            # Drop rows with missing product names
            df_final.dropna(subset=['Product_Name'], inplace=True)

            # Save to Excel
            with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
                sheet_name = FILE_NAME.replace('.xlsx', '')[:31]
                df_final.to_excel(writer, sheet_name=sheet_name, index=False)

            print(f"\n✅ Final product data saved to: {FILE_NAME}")

        else:
            print("⚠️ No products found at all.")

    finally:
        driver.quit()
        print("Browser closed.")

if __name__ == "__main__":
    main()
