import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

# Specify the path to the chromedriver executable
driver_service = Service(executable_path='C:\Program Files (x86)\chromedriver.exe')
# driver_service = Service(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=driver_service)

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(r'C:\Users\semja\Documents\Python\ASINS_weblink.xlsx', header=None, names=["url"])

# Create a list to store the data
data = []

# Iterate over the list of URLs
for url in df["url"]:
    url = f"https://www.amazon.de/dp/{url}"
    driver.get(url)

    # Wait for the page to load
    try:
        wait = WebDriverWait(driver, 2)
        if driver.find_element(By.ID, "buy-now-button"):
            element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tabular-buybox-container")))
            # Find elements
            sold_by = driver.find_element(By.ID, "sellerProfileTriggerId")
            sold_by = driver.execute_script("return arguments[0].innerHTML;", sold_by)
            price = driver.find_element(By.CLASS_NAME, "a-offscreen")
            price = driver.execute_script("return arguments[0].innerHTML;", price)
            delivery_price = driver.find_element(By.XPATH, "//span[@data-csa-c-delivery-price]")
            delivery_price = delivery_price.get_attribute('data-csa-c-delivery-price')
            delivery_time = driver.find_element(By.CSS_SELECTOR, 'span[data-csa-c-delivery-time]')
            delivery_time = delivery_time.get_attribute('data-csa-c-delivery-time')
            data.append([url, sold_by, price, delivery_price, delivery_time])
            continue

    except Exception:
        print(f"Exception: {url}")
        continue

# Create a pandas DataFrame from the data list
result_df = pd.DataFrame(data, columns=["url", "Sold By", "Price", "Delivery Price", "Delivery Time"])

# Define the CSV file path
csv_file_path = r'C:\Users\semja\Documents\Python\Buybox_Amazon\data.csv'

# Export DataFrame to CSV
result_df.to_csv(csv_file_path, index=False)
