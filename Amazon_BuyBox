import re

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from tqdm import tqdm


# Function to extract the price from the text
def extract_float(text):
    match = re.search(r'(\d+,\d+)', text)
    if match:
        return match.group(1)
    return None


# Function to extract the price from the text
def extract_int(text):
    match = re.search(r'(\d+)', text)
    if match:
        return match.group(1)
    return None


def scrape_amazon_data(url, driver):
    try:
        driver.get(url)

        # Wait for the 'buy-now-button' to be clickable
        buy_now_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, 'buyNow')))

        # Scroll to the buy-now-button to ensure it's in the viewport
        driver.execute_script("arguments[0].scrollIntoView();", buy_now_button)

        try:
            try:
                sold_by = driver.execute_script('return document.querySelector("#sellerProfileTriggerId").innerText')
            except Exception:
                try:
                    sold_by = driver.execute_script(
                        'return document.querySelector("#merchantInfoFeature_feature_div > div.offer-display-feature-text > div > span").innerText')
                except Exception:
                    try:
                        sold_by = driver.find_element(By.XPATH,
                                                      '//*[@id="merchantInfoFeature_feature_div"]/div[2]/div/span')
                    except NoSuchElementException:
                        try:
                            sold_by = driver.find_element(By.ID, "sellerProfileTriggerId")
                        except NoSuchElementException:
                            sold_by = "none"
        except Exception:
            sold_by = ""

        try:
            min_price_element1 = driver.execute_script(
                'return document.querySelector("#olpLinkWidget_feature_div > div.a-section.olp-link-widget > span > a > div > div > span.a-price > span:nth-child(2)")')
            min_price_int = extract_int(min_price_element1.text.strip())
            min_price_element2 = driver.execute_script(
                'return document.querySelector("#olpLinkWidget_feature_div > div.a-section.olp-link-widget > span > a > div > div > span.a-price > span:nth-child(2) > span.a-price-fraction")')
            min_price_dec = extract_int(min_price_element2.text.strip())
            min_price = min_price_int + "," + min_price_dec
        except Exception:
            min_price = "none"

        try:
            try:
                delivery_price_element = driver.find_element(By.XPATH,
                                                             '//*[@id="mir-layout-DELIVERY_BLOCK-slot-PRIMARY_DELIVERY_MESSAGE_LARGE"]/span')
                delivery_price = extract_float(
                    delivery_price_element.text.strip())  # Trim any leading/trailing whitespace
            except NoSuchElementException:
                delivery_price = "none"
        except Exception:
            delivery_price = ""

        try:
            try:
                delivery_time = driver.find_element(By.CSS_SELECTOR, 'span[data-csa-c-delivery-time]').get_attribute(
                    'data-csa-c-delivery-time')
            except NoSuchElementException:
                delivery_time = "none"
        except Exception:
            delivery_time = ""

        # Use JavaScript to extract price
        try:
            price_js = driver.execute_script('return document.querySelector("#corePrice_feature_div").innerText')
            price = extract_float(price_js) if price_js else ""
        except Exception:
            price = ""

        try:
            fulfilled_by = driver.execute_script(
                'return document.querySelector("#fulfillerInfoFeature_feature_div > div.offer-display-feature-text > div > span").innerText')
        except Exception:
            fulfilled_by = ""

        try:
            try:
                availability_element = driver.find_element(By.XPATH, '//*[@id="availability"]/span')
                availability = extract_int(availability_element.text)
            except NoSuchElementException:
                availability = ""
        except Exception:
            availability = ""

        try:
            reviews_js = driver.execute_script(
                'return document.querySelector("#acrPopover > span.a-declarative > a > span").innerText')
            reviews = reviews_js.strip() if reviews_js else ""
        except Exception:
            reviews = ""

        try:
            try:
                count_reviews = driver.find_element(By.XPATH,
                                                    '//*[@id="cm_cr_dp_d_rating_histogram"]/div[3]/span').text.strip()
            except NoSuchElementException:
                count_reviews = "none"
        except Exception:
            count_reviews = ""

        try:
            try:
                bestseller_rank_element1 = driver.find_element(By.XPATH,
                                                               '//*[@id="productDetails_detailBullets_sections1"]/tbody/tr[3]/td/span')
                bestseller_rank = bestseller_rank_element1.text.replace('\n', ' | ').strip()
            except NoSuchElementException:
                try:
                    bestseller_rank_element2 = driver.find_element(By.XPATH,
                                                                   '//*[@id="detailBulletsWrapper_feature_div"]/ul[1]/li/span')
                    bestseller_rank = bestseller_rank_element2.text.replace('\n', ' | ').strip()
                except NoSuchElementException:
                    bestseller_rank = "none"
        except Exception:
            bestseller_rank = ""

        return [url, sold_by, fulfilled_by, price, delivery_price, delivery_time, min_price, availability, reviews,
                count_reviews, bestseller_rank]

    except Exception as e:
        print(f"Error: {str(e)} occurred for URL: {url}")
        return [url, "", "", "", "", "", "", "", "", "", ""]


def main():
    # Specify the path to the chromedriver executable
    chrome_driver_path = r"C:\Program Files (x86)\chromedriver.exe"

    # Set up Chrome options and specify the executable path
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument(f"executable_path={chrome_driver_path}")

    # Create the driver using Chrome options
    driver = webdriver.Chrome(options=chrome_options)

    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(r'C:\Users\semja\Documents\Python\ASINS_weblink.xlsx', header=None, names=["url"])

    # Create a list to store the data
    data = []

    # Get the total number of URLs to process
    total_urls = len(df["url"])

    # Use tqdm to create a progress bar
    with tqdm(total=total_urls, desc="Progress") as pbar:
        for url in df["url"]:
            url = f"https://www.amazon.de/dp/{url}"
            scraped_data = scrape_amazon_data(url, driver)
            data.append(scraped_data)
            pbar.update(1)  # Update progress bar

    # Create a pandas DataFrame from the data list
    result_df = pd.DataFrame(data,
                             columns=["url", "Sold By", "fulfilled_by", "Price", "Delivery Price", "Delivery Time",
                                      "min_price", "availability", "reviews", "count_reviews",
                                      "bestseller_rank"])

    # Define the CSV file path
    csv_file_path = r'C:\Users\semja\Documents\Python\Buybox_Amazon\data.csv'

    # Export DataFrame to CSV
    result_df.to_csv(csv_file_path, index=False)

    # Close the WebDriver
    driver.quit()


if __name__ == "__main__":
    main()
