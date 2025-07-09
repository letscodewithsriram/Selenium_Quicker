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
with open("creds.json","r") as credshr:
    creds = json.load(credshr)

def convert_to_excel(perf_info):
    print(perf_info)
    workbook = xlsxwriter.workbook("Monthly_Stats.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Name')
    worksheet.write('B1', 'Support Tickets')
    worksheet.write('C1', 'Escalated Tickets')

    workbook.close()
def get_tkt_timer_info(url):
    print(url)
    driver.get(url)
    timer_tbl = driver.find_element(By.ID, "Table8")
    timer_tblrows = timer_tbl.find_elements(By.TAG_NAME, "tr")

    for row in timer_tblrows:
        cols = row.find_elements(By.TAG_NAME, "td")
        data = [column.text.strip() for column in cols]
        print(data)
    return None

## FSR Credentials
username = creds["username"]
password = creds["password"]
l3team = creds["l3team"]

### Setup Chrome Options
chrome_options = Options()
# chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-extensions')
chrome_options.accept_insecure_certs = True

driver = webdriver.Chrome(options=chrome_options)

driver.get("https://forticare.fortinet.com")
time.sleep(2)
driver.maximize_window()
# <input id="id_username_input" type="text" placeholder="Username" name="username" value="">
username_field = driver.find_element(By.ID, "id_username_input")
# <input id="id_submit_btn" type="submit" class="submit" value="Next">
next_btn = driver.find_element(By.ID, "id_submit_btn")
username_field.send_keys(username)
next_btn.click()
time.sleep(5)
# <input id="id_password" type="password" name="password" placeholder="Password">
password_field = driver.find_element(By.ID, "id_password")
password_field.send_keys(password)
# <button type="submit" class="submit" value="Login">Login</button>
login_btn = driver.find_element(By.XPATH, "//button[text()='Login']")
login_btn.click()
time.sleep(30)
#!# token_code = input("Enter the FortiToken Number: ")
#!# print("You've entered: ", token_code)
# <input type="text" name="token_code" id="id_password">
#!# token_field = driver.find_element(By.NAME, "token_code")
#!# token_field.send_keys(token_code)
# <input type="submit" class="submit" value="GO">
#!# submit_btn = driver.find_element(By.CLASS_NAME, "submit")
#!# submit_btn.click()
#!# time.sleep(5)

# <a href="/CustomerSupport/SupportTeam/SearchTicketPr.aspx">Search</a>

l3_team_efforts = {}

for employee in l3team:
    employee = employee.strip()
    print(employee)
    l3_team_efforts[employee] = []
    print(l3_team_efforts)

    driver.get("https://forticare.fortinet.com/CustomerSupport/SupportTeam/SearchTicketPr.aspx")

    closeFromDate = driver.find_element(By.ID, "ctl00_MainContent_SearchTickets_TB_CloseFromDate")
    closeFromDate.clear()
    closeToDate = driver.find_element(By.ID,"ctl00_MainContent_SearchTickets_TB_CloseToDate")
    closeToDate.clear()
    # <input name="ctl00$MainContent$SearchTickets$TB_CloseFromDate" type="text" value="MM/DD/YYYY" id="ctl00_MainContent_SearchTickets_TB_CloseFromDate" class="searchlabel" style="width:108px;">
    # <input name="ctl00$MainContent$SearchTickets$TB_CloseToDate" type="text" value="MM/DD/YYYY" id="ctl00_MainContent_SearchTickets_TB_CloseToDate" class="searchlabel" style="width:117px;">
    closeFromDate.send_keys("04/01/2025")
    closeToDate.send_keys("06/30/2025")
    time.sleep(5)

    status_dd = driver.find_element(By.ID, "ctl00_MainContent_SearchTickets_DDL_TicketStatus")
    select = Select(status_dd)
    select.select_by_visible_text("Closed")

    ## <input name="ctl00$MainContent$SearchTickets$TB_KeyWords"
    # type="text" id="ctl00_MainContent_SearchTickets_TB_KeyWords" style="width:426px;">

    ticketComments = driver.find_element(By.ID, "ctl00_MainContent_SearchTickets_TB_KeyWords")
    ticketComments.clear()
    ticketComments.send_keys("_" + employee + "_")
    time.sleep(5)

    # category_dd = driver.find_element(By.ID, "ctl00_MainContent_SearchTickets_DDL_Category_DDL_ProductType")
    # select = Select(category_dd)
    # select.select_by_visible_text("FortiSOAR")
    # time.sleep(5)
    #
    # queue_dd = driver.find_element(By.ID, "ctl00_MainContent_SearchTickets_DDL_Queue_Name")
    # select = Select(queue_dd)
    # select.select_by_visible_text("AMER_SOAR")
    # time.sleep(5)
    #

    ## <input type="submit" name="ctl00$MainContent$SearchTickets$B_Search" value="Search" id="ctl00_MainContent_SearchTickets_B_Search">
    search_btn = driver.find_element(By.ID, "ctl00_MainContent_SearchTickets_B_Search")
    search_btn.click()
    time.sleep(5)

    # pagesize_dd = driver.find_element(By.ID, "ctl00_MainContent_DDL_PageSize")
    # select = Select(pagesize_dd)
    # select.select_by_visible_text("800")
    # time.sleep(5)

    try:
        ticket_tbl = driver.find_element(By.ID, "ctl00_MainContent_DG_TicketList")
        ticket_tblrows = ticket_tbl.find_elements(By.TAG_NAME, "tr")

        for row in ticket_tblrows:
            cols = row.find_elements(By.TAG_NAME, "td")
            data = [column.text.strip() for column in cols]
            ## https://forticare.fortinet.com/customersupport/default.aspx?TID=' + ticketId
            if data[0] != "#":
                if int(data[0]) > 1000:
                    tid = data[0]
                    towner = data[6]
                    print(data)
                    l3_team_efforts[employee].append(str(tid) + "~" + str(towner))
                    # tid_url = "https://forticare.fortinet.com/customersupport/default.aspx?TID=" + str(tid)
                    # print(tid_url)
    except NoSuchElementException:
        print("Table Not Found")
time.sleep(5)

print(l3_team_efforts)
convert_to_excel(l3_team_efforts)
driver.quit()