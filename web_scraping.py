from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import random

# Initialize ChromeDriver
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
start_date_input.send_keys("1403/08/20")

time.sleep(random.uniform(1, 3))
end_date_input = wait.until(EC.visibility_of_element_located((By.NAME, 'ctl05$ReportsHeaderControl$ToDate')))
end_date_input.clear()
end_date_input.send_keys("1403/08/23")

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

# Extract headers
columns = []
header_cells = driver.find_elements(By.CSS_SELECTOR, "#AmareMoamelatGridTbl thead th")
for header in header_cells:
    header_text = header.get_attribute("innerText").strip() or header.get_attribute("innerHTML").strip()
    if header_text:
        columns.append(header_text)

data = []
for row in rows:
    cells = row.find_elements(By.TAG_NAME, "td")
    cell_data = []
    for cell in cells:
        # Replace unwanted HTML elements with "N/A" or clean them
        text = cell.get_attribute("innerText").replace('<i class="fa fa-exclamation"></i>', 'N/A').strip()
        cell_data.append(text)

    # Add the data from this page to all_data
    data.append(cell_data)

# Save to CSV
df = pd.DataFrame(data, columns=columns)
df.to_csv("ime_transaction_data.csv", index=False, encoding="utf-8-sig")

print("Data saved to ime_transaction_data.csv")
driver.quit()
