from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import os
from datetime import datetime, timedelta
import pandas as pd
import random
from selenium.webdriver.chrome.service import Service

# Function to check if the file has been downloaded
def wait_for_download(download_path, timeout=20):
    """
    Wait for a file to be downloaded within a specified timeout.
    :param download_path: Path to the download directory
    :param timeout: Maximum time to wait for the download (in seconds)
    :return: The name of the downloaded file if successful, None otherwise
    """
    end_time = time.time() + timeout
    while time.time() < end_time:
        # Check if any file in the directory ends with '.json'
        for filename in os.listdir(download_path):
            if filename.endswith('.json'):
                return filename
        time.sleep(1)  # Wait for 1 second before checking again
    return None

# Define dates
date_source = pd.read_excel("DimDate.xlsx")
dates_list = date_source["time_Title_Year_Month_Day"].values.tolist()
dates = []
count = 19
# Create date ranges by splitting the list into intervals of 20 days
for i in range(0, len(dates_list) - 20, 20):
    dates.append([dates_list[i], dates_list[count]])
    count += 20

# Loop through each date range in the list
for row, date in enumerate(dates):
    # 2: Initialize ChromeDriver
    service = Service(r'C:\chromedriver-win64\chromedriver.exe')
    driver = webdriver.Chrome(service=service)

    # Open the target page
    url = "https://www.ime.co.ir/offer-stat.html"
    driver.get(url)

    wait = WebDriverWait(driver, 15)

    # Interact with the dropdown to select all columns
    dropdown_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-toggle="dropdown"]')))
    dropdown_button.click()

    checkboxes = driver.find_elements(By.CSS_SELECTOR, 'ul.dropdown-menu input[type="checkbox"]')
    for checkbox in checkboxes:
        if not checkbox.is_selected():
            checkbox.click()
            time.sleep(random.uniform(.2, .5))

    # Set start and end dates
    time.sleep(random.uniform(2, 4))
    start_date_input = wait.until(EC.visibility_of_element_located((By.NAME, 'ctl05$ReportsHeaderControl$FromDate')))
    start_date_input.clear()
    start_date_input.send_keys(date[0])

    time.sleep(random.uniform(1, 3))
    end_date_input = wait.until(EC.visibility_of_element_located((By.NAME, 'ctl05$ReportsHeaderControl$ToDate')))
    end_date_input.clear()
    end_date_input.send_keys(date[1])

    time.sleep(random.uniform(1, 3))
    # Click the "نمایش" button
    display_button = driver.find_element(By.ID, "FillGrid")
    display_button.click()

    # Wait for the table to load
    time.sleep(5)
    # Wait for the table to load
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#AmareMoamelatGridTbl tbody tr:last-child")))

    # Scroll until all rows are loaded
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

    # Wait for the table to refresh
    time.sleep(3)  # Adjust this sleep time if necessary

    # Click on the dropdown button to open export options
    try:
        dropdown_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@class='btn btn-default dropdown-toggle' and @title='Export data']")))
        dropdown_button.click()  # Click the dropdown to show export options

        # Click on the JSON option in the dropdown menu
        json_download_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-type='json']/a")))
        json_download_link.click()  # Click the JSON link to download the data
    except TimeoutException:
        print("Failed to find the dropdown button or JSON link. Skipping this date range.")
        driver.quit()
        continue

    # Wait for the download to complete
    download_path = r"C:\Users\Mosalas\Downloads"  # Change to your download path
    downloaded_file = wait_for_download(download_path)
    if downloaded_file:
        # Rename the downloaded file to include the date range
        date_1 = date[0].replace("/", "_")
        date_2 = date[1].replace("/", "_")
        new_file_name = f"data_{date_1 }_to_{date_2}.json"
        os.rename(os.path.join(download_path, downloaded_file), os.path.join(download_path, new_file_name))
        print(f"Downloaded and renamed data for {date_1} to {date_2} to {new_file_name}")
    else:
        print(f"Failed to download data for {date[0]} to {date[1]}")

    # Close the browser
    driver.quit()

print("Data extraction completed successfully.")
