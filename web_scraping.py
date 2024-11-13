import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import random



# Initialize undetected ChromeDriver to bypass detection
# options = uc.ChromeOptions()
# options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36")
# driver = uc.Chrome(options=options)

# Initialize ChromeDriver
service = Service(r'C:\Users\sohrabi.s\Documents\chromedriver-win64\chromedriver.exe')
driver = webdriver.Chrome(service=service)

# Open the target page
url = "https://www.ime.co.ir/offer-stat.html"
driver.get(url)

# Wait for the elements to load
wait = WebDriverWait(driver, 10)

# Set start and end dates with random delays
time.sleep(random.uniform(1, 3))
start_date_input = wait.until(EC.visibility_of_element_located((By.NAME, 'ctl05$ReportsHeaderControl$FromDate')))
start_date_input.clear()
start_date_input.send_keys("1403/08/20")

time.sleep(random.uniform(1, 3))
end_date_input = wait.until(EC.visibility_of_element_located((By.NAME, 'ctl05$ReportsHeaderControl$ToDate')))
end_date_input.clear()
end_date_input.send_keys("1403/08/23")

# Click the "نمایش" button with a random delay
time.sleep(random.uniform(1, 3))
display_button = driver.find_element(By.ID, "FillGrid")
display_button.click()

# Wait for the data table to load
time.sleep(5)  # or adjust with WebDriverWait if necessary

# Extract column names dynamically from the table header
columns = []
header_cells = driver.find_elements(By.CSS_SELECTOR, "table thead th")
for header in header_cells:
    columns.append(header.text.strip())
print(columns)

# Extract data from table rows
data = []
rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
for row in rows:
    cells = row.find_elements(By.TAG_NAME, "td")
    cell_data = [cell.text.strip() for cell in cells]
    if cell_data:
        data.append(cell_data)
print(data)
# Close the browser
driver.quit()
print('finished')

# Convert to DataFrame and save to CSV
# df = pd.DataFrame(data, columns=columns)
# df.to_csv("ime_transaction_data.csv", index=False, encoding="utf-8-sig")
# print("Data saved to ime_transaction_data.csv")