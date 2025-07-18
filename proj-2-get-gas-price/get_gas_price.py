from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import crypto
import time
import cmaths
import json
import xlsxwriter
import re
from datetime import date
from dateutil.relativedelta import relativedelta
import csv
chrome_options = Options()
# chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-extensions')
chrome_options.accept_insecure_certs = True

driver = webdriver.Chrome(options=chrome_options)

driver.get("https://www.cbc.ca/bc/gasprices/")
time.sleep(2)
driver.maximize_window()

table = driver.find_element(By.ID, "prices")
print("Table found!")
cols = table.find_elements(By.TAG_NAME, "tr")
parsed_lines = []
for row in cols:
    # print(row)
    data = row.text.strip()
    data = ' '.join(data.split())
    # print(data)
    # Regex pattern: (price) (brand name) (address) (date/time)
    pattern = re.compile(r"^(\d+\.\d+)\s+(.+?)\s+(\d{3,5} .+?)\s+(Jul \d{1,2}, \d{1,2}:\d{2} [AP]M)$")
    match = pattern.match(data)
    if match:
        price, brand, address, timestamp = match.groups()
        parsed_lines.append(f"{price}${brand}${address}${timestamp}")
print(parsed_lines)
with open("gas_prices.csv", "a", newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)
    # writer.writerow(["Price", "Brand", "Address", "Timestamp"])  # Header
    for entry in parsed_lines:
        parts = entry.split("$")  # Only split into 4 parts (limit in case of colons in address)
        writer.writerow(parts)

time.sleep(5)