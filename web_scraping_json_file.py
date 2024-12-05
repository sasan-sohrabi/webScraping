from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time
import os
import pandas as pd
import random


# Function to wait for and find the last downloaded file
def wait_for_download(download_path, initial_files, timeout=30, check_interval=1):
    """
    Wait for a new file to appear in the download directory.
    :param download_path: Path to the download directory
    :param timeout: Maximum time to wait for the download (in seconds)
    :param check_interval: How often to check for a new file (in seconds)
    :return: The full path of the downloaded file, or None if timeout is reached
    """
    end_time = time.time() + timeout

    while time.time() < end_time:
        current_files = set(os.listdir(download_path))
        new_files = current_files - initial_files  # Identify newly added files
        if new_files:
            return os.path.join(download_path, new_files.pop())
        time.sleep(check_interval)

    return None  # Timeout reached, no file detected


# Define dates
date_source = pd.read_excel("DimDate.xlsx")
dates_list = date_source["time_Title_Year_Month_Day"].values.tolist()
dates = []
count = 9
for i in range(0, len(dates_list) - 10, 10):
    dates.append([dates_list[i], dates_list[count]])
    count += 10

# Define download path and ChromeDriver service
download_path = r"C:\Users\Mosalas\Downloads"
chrome_service = Service(r'C:\chromedriver-win64\chromedriver.exe')

# Iterate through each date range
for date_range in dates:
    # Initialize the WebDriver
    driver = webdriver.Chrome(service=chrome_service)
    driver.get("https://www.ime.co.ir/offer-stat.html")
    initial_files = set(os.listdir(download_path))  # Files before download starts

    wait = WebDriverWait(driver, 15)

    try:
        # Select all columns
        dropdown_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-toggle="dropdown"]')))
        dropdown_button.click()

        checkboxes = driver.find_elements(By.CSS_SELECTOR, 'ul.dropdown-menu input[type="checkbox"]')
        for checkbox in checkboxes:
            if not checkbox.is_selected():
                checkbox.click()
                time.sleep(random.uniform(0.2, 0.5))

        # Enter start and end dates
        start_date_input = wait.until(
            EC.visibility_of_element_located((By.NAME, 'ctl05$ReportsHeaderControl$FromDate')))
        start_date_input.clear()
        start_date_input.send_keys(date_range[0])

        end_date_input = wait.until(EC.visibility_of_element_located((By.NAME, 'ctl05$ReportsHeaderControl$ToDate')))
        end_date_input.clear()
        end_date_input.send_keys(date_range[1])

        # Click the "نمایش" button to load the table
        display_button = driver.find_element(By.ID, "FillGrid")
        display_button.click()

        # Wait for the table to load and scroll to load all rows
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#AmareMoamelatGridTbl tbody tr:last-child")))
        time.sleep(random.uniform(20, 30))
        last_row_count = 0
        while True:
            rows = driver.find_elements(By.CSS_SELECTOR, "#AmareMoamelatGridTbl tbody tr")
            if len(rows) > last_row_count:
                last_row_count = len(rows)
                driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight",
                                      driver.find_element(By.ID, "AmareMoamelatGridTbl"))
                time.sleep(2)
            else:
                break

        # Export data as JSON
        export_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Export data']")))
        export_dropdown.click()

        json_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-type='json']/a")))
        json_option.click()

        # Wait for download and rename the file
        downloaded_file = wait_for_download(download_path, initial_files)
        if downloaded_file:
            date_1 = date_range[0].replace("/", "_")
            date_2 = date_range[1].replace("/", "_")
            new_file_name = f"data_{date_1}_to_{date_2}.json"
            os.rename(downloaded_file, os.path.join(download_path, new_file_name))
            print(f"Renamed downloaded file to: {new_file_name}")
        else:
            print(f"Download failed for date range: {date_range[0]} to {date_range[1]}")

    except Exception as e:
        print(f"Error processing date range {date_range[0]} to {date_range[1]}: {e}")

    finally:
        # Close the browser
        driver.quit()

print("Data extraction completed successfully.")
