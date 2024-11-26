import pyodbc
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random
import pandas as pd

# 1: Connect to SQL Server
conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=.\SASANENERGY;"  
    "DATABASE=IME_WebScraping;"  
    "Trusted_Connection=yes;"
)
cursor = conn.cursor()

# Define dates
date_source = pd.read_excel("DimDate.xlsx")
dates_list = date_source["time_Title_Year_Month_Day"].values.tolist()
dates = []
count = 19
for i in range(0, len(dates_list) - 20, 20):
    dates.append([dates_list[i], dates_list[count]])
    count += 20
#dates = [["1403/08/01", "1403/08/07"], ["1403/08/08", "1403/08/15"], ["1403/08/16", "1403/08/23"], ["1403/08/23", "1403/08/30"]]


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
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", driver.find_element(By.ID, "AmareMoamelatGridTbl"))
            time.sleep(2)
        else:
            break

    # Click the page-size dropdown button to display options
    page_size_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span.page-list button")))
    page_size_button.click()

    # Select the "100" option from the dropdown
    option_100 = wait.until(EC.element_to_be_clickable((By.XPATH, "//ul[@class='dropdown-menu']//a[text()='100']")))
    option_100.click()
    time.sleep(3)  # Wait for the table to reload with 100 rows

    # 3: Extract headers and Create table on database
    columns = []
    header_cells = driver.find_elements(By.CSS_SELECTOR, "#AmareMoamelatGridTbl thead th")
    for header in header_cells:
        header_text = header.get_attribute("innerText").strip() or header.get_attribute("innerHTML").strip()
        if header_text:
            columns.append(header_text)

    if row == 0:
        # Define table creation query
        table_name = "IME_Transactions"
        create_table_query = f"CREATE TABLE {table_name} (ID INT IDENTITY(1,1) PRIMARY KEY, "
        for column in columns:
            create_table_query += f"[{column}] NVARCHAR(MAX), "
        create_table_query = create_table_query.rstrip(", ") + ");"

        # Execute table creation
        try:
            cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
            cursor.execute(create_table_query)
            conn.commit()
            print(f"Table '{table_name}' created successfully.")
        except Exception as e:
            print("Error creating table:", e)

    # 4: Loop through each page and insert data into SQL Server
    last_page = int(driver.find_element(By.CSS_SELECTOR, "li.page-last a").text.strip())
    # data = []
    page = 0
    while True:
        page += 1
        print(page)
        try:
            # Re-fetch rows to avoid stale element reference
            rows = driver.find_elements(By.CSS_SELECTOR, "#AmareMoamelatGridTbl tbody tr")
            for i in range(len(rows)):
                rows = driver.find_elements(By.CSS_SELECTOR, "#AmareMoamelatGridTbl tbody tr")  # Refresh rows
                cells = rows[i].find_elements(By.TAG_NAME, "td")
                cell_data = [cell.get_attribute("innerText").strip() for cell in cells]
                # data.append(cell_data)

                placeholders = ", ".join(["?"] * len(columns))
                insert_query = f"INSERT INTO {table_name} ({', '.join(['[' + col + ']' for col in columns])}) VALUES ({placeholders})"
                cursor.execute(insert_query, cell_data)

            # Commit after each page
            conn.commit()

        except:
            print("Encountered a stale element, retrying...")
            continue

        # Navigate to the next page if not the last
        if page < last_page:
            next_page = driver.find_element(By.CSS_SELECTOR, "li.page-next a")
            next_page.click()
            print("clicked next page")
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#AmareMoamelatGridTbl tbody tr"))
            )
        else:
            break
    driver.quit()


# Save to CSV
# df = pd.DataFrame(data, columns=columns)
# df.to_csv("ime_transaction_data.csv", index=False, encoding="utf-8-sig")
# print("Data saved to ime_transaction_data.csv")

# Close SQL Server connection and WebDriver
cursor.close()
conn.close()
driver.quit()

print("Data extraction and insertion completed successfully.")
