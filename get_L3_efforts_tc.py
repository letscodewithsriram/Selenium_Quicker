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

today = date.today()
last_month = today - relativedelta(months=1)
formatted_last_month = last_month.strftime('%b-%Y')  # 'Jun-2025'
formatted_xls_last_month = last_month.strftime('%b_%Y')  # 'Jun-2025'

with open("creds.json","r") as credshr:
    creds = json.load(credshr)

def convert_to_excel(perf_info):
    print(perf_info)
    perf_info = {'SRIRAM': ['10787968~Farhan Ahmed', '10662988~Lovepreet Singh', '10618208~Ankit Jain', '10607081~Sriram Ramanujam', '10568171~Prachi Patel',
                            '10556936~Himanshu Jain', '10506535~Ankit Jain', '10335074~Sriram Ramanujam'],
                 'PHIL': ['10834149~Harsh Patel', '10808630~Sagar Ajudiya', '10808040~Andrei Portnov', '10781682~Santhosh Shanmugam', '10772826~Ahmad Srour',
                          '10764261~Ahmad Srour', '10757407~Harsh Patel', '10666571~Andrei Portnov', '10611309~Harsh Patel', '10611024~Ahmad Srour', '10610281~Harsh Patel',
                          '10592786~Phil Rose', '10581257~Ahmad Srour', '10179306~Ahmad Srour', '9093372~Saif Ali Momin'],
                 'LOVEPREET': ['10750191~Harsh Patel', '10710545~Harsh Patel', '10707139~Harsh Patel', '10702937~Harsh Patel',
                               '10631843~Harsh Patel', '10604987~Santhosh Shanmugam', '10602581~Sriram Ramanujam', '10600119~Ahmad Srour', '10581097~Sangeet Sharma',
                               '10571623~Lovepreet Singh', '10542560~Antonio Elves Alves Ribeiro', '10533232~Selvam Sitaraman', '10482385~Adrian Hill', '10471510~Selvam Sitaraman',
                               '10470732~Ahmad Srour', '10388286~Tarranum Bano', '10236254~Daniel Zhong', '9889870~Antonio Elves Alves Ribeiro'],
                 'VANITA': ['10618811~Muhammad Amer Rashid', '10618811~Muhammad Amer Rashid','10544944~Sriram Ramanujam', '10479457~Santhosh Shanmugam', '10467627~Vanita Fernandes', '10403057~Sriram Ramanujam',
                            '10167866~Sriram Ramanujam'],
                 'ESC_SRIRAM': [],
                 'ESC_PHIL': ['10706806~Phil Rose', '10615420~Phil Rose', '10605550~Phil Rose', '10592786~Phil Rose', '10446718~Phil Rose', '10407446~Phil Rose'],
                 'ESC_LOVEPREET': ['10669888~Lovepreet Singh', '10600635~Lovepreet Singh', '10571623~Lovepreet Singh', '10527249~Lovepreet Singh', '10317933~Lovepreet Singh'],
                 'ESC_VANITA': ['10699878~Vanita Fernandes', '10621071~Vanita Fernandes', '10518576~Vanita Fernandes', '10467627~Vanita Fernandes',
                                '10457837~Sangeet Sharma', '10298672~Vanita Fernandes', '9002707~Vanita Fernandes'],
                 'DL': ['10762155~Kavithas Thevarajah', '10762155~Kavithas Thevarajah', '10710191~Harsh Patel', '10110545~Harsh Patel', '10701139~Harsh Patel', '10712937~Harsh Patel',
                               '10631843~Harsh Patel', '10604987~Santhosh Shanmugam', '10602581~Sriram Ramanujam', '10600119~Ahmad Srour', '10581097~Sangeet Sharma',
                               '10571623~Lovepreet Singh', '10542560~Antonio Elves Alves Ribeiro', '10533232~Selvam Sitaraman', '10482385~Adrian Hill', '10471510~Selvam Sitaraman',
                               '10470732~Ahmad Srour', '10388286~Tarranum Bano', '10236254~Daniel Zhong', '9889870~Antonio Elves Alves Ribeiro']}
    print(perf_info)
    print(type(perf_info))

    # merge_assist_escalated

    for l3mem in creds["l3team"]:
        if not re.search("ESC", l3mem) and not re.search("DL", l3mem):
            print(l3mem)
            # To check two tags (both assist and escalated) in tickets to avoid double counting.
            for esctkt in perf_info["ESC_" + l3mem]:
                print(esctkt.strip().split('~')[0])
                for asttkts in perf_info[l3mem]:
                    if re.search(esctkt.strip().split('~')[0], asttkts):
                        perf_info[l3mem].remove(asttkts)
                        print(asttkts)

    scores = []

    for k,v in perf_info.items():
        if not re.search("ESC", k) and not re.search("DL", k):
            scores.append((k, len(v)*0.25 + len(perf_info["ESC_" + k]) + creds['training'][k]*0.5))

    max_score = max(score for name, score in scores)
    toppers = [name for name, score in scores if score == max_score]

    workbook = xlsxwriter.Workbook("L3_Monthly_Stat_" + str(formatted_xls_last_month) + ".xlsx")

    cell_format = workbook.add_format({
        'bold': False,
        'font_color': 'black',
        'font_size': 11,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'text_wrap': True
    })

    table_header_format = workbook.add_format({
        'bold': True,
        'font_color': 'black',
        'bg_color': 'red',
        'font_size': 11,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    topper_format = workbook.add_format({
        'bold': True,
        'font_color': 'black',
        'bg_color': 'green',
        'font_size': 11,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    header_format = workbook.add_format({
        'bold': True,
        'font_color': 'black',
        'bg_color': 'white',
        'font_size': 20,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    sub_header_format = workbook.add_format({
        'bold': True,
        'font_color': 'black',
        'bg_color': 'white',
        'font_size': 14,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    worksheet = workbook.add_worksheet()
    worksheet.set_column("A:A", 20, cell_format)
    worksheet.set_column("B:B", 14, cell_format)
    worksheet.set_column("C:C", 14, cell_format)
    worksheet.set_column("D:D", 14, cell_format)
    worksheet.set_column("E:E", 14, cell_format)
    worksheet.set_column("F:F", 35, cell_format)
    worksheet.set_column("G:G", 35, cell_format)
    worksheet.insert_image('A1', 'Fortinet-logomark-rgb-red.png', {'x_scale': 0.75, 'y_scale': 0.75, 'x_offset':30, 'y_offset':10})
    worksheet.insert_image('G1', 'Fortinet-logo-rgb-black-red.png', {'x_scale': 1.25, 'y_scale': 1.25, 'x_offset':5, 'y_offset':15})
    merge_format = workbook.add_format({'align': 'center', 'bold': True})

    worksheet.merge_range('B2:F2', "Generated On - " + str(today) + "[" + str(creds['startdate']) + " - " + str(creds['enddate'])+ "]", sub_header_format)
    worksheet.merge_range('B1:F1', "Monthly Report : " + str(formatted_last_month), header_format)
    worksheet.merge_range('B3:F3', "Queues Covered: AMER_SOAR, AMER_FMG_FAZ", sub_header_format)
    worksheet.merge_range('A1:A3', "")
    worksheet.merge_range('G1:G3', "")

    row = 3

    worksheet.write(row, 0, 'NAME', table_header_format)
    worksheet.write(row, 1, 'SCORES', table_header_format)
    worksheet.write(row, 2, 'ASSIST', table_header_format)
    worksheet.write(row, 3, 'ESCALATED', table_header_format)
    worksheet.write(row, 4, 'TOPICS', table_header_format)
    worksheet.write(row, 5, "ASSIST TICKETS", table_header_format)
    worksheet.write(row, 6, "ESCALATED TICKETS", table_header_format)

    row = row + 1
    for k,v in perf_info.items():
        col = 0
        temp = []
        ## Assist
        if not re.search("ESC", k) and not re.search("DL", k):
            worksheet.write(row, col, k)
            # worksheet.write(row, col + 1, len(v)*creds['scoring_factor']['assist'] + len(perf_info["ESC_" + k])*creds['scoring_factor']['escalate'])
            if k in toppers:
                worksheet.write(row, col + 1, len(v)*creds['scoring_factor']['assist'] + len(perf_info["ESC_" + k])*creds['scoring_factor']['escalate'] + creds['training'][k]*creds['scoring_factor']['topic'], topper_format)
            else:
                worksheet.write(row, col + 1, len(v)*creds['scoring_factor']['assist'] + len(perf_info["ESC_" + k])*creds['scoring_factor']['escalate']  + creds['training'][k]*creds['scoring_factor']['topic'])
            worksheet.write(row, col + 2, len(v))
            worksheet.write(row, col + 3, len(perf_info["ESC_" + k]))
            worksheet.write(row, col + 4, creds['training'][k])
            worksheet.write(row, col + 5, ''.join(tk.strip().split("~")[0] + ' ' for tk in v))
            worksheet.write(row, col + 6, ''.join(tk.strip().split("~")[0] + ' ' for tk in perf_info["ESC_" + k]))
        row = row + 1

    workbook.close()

    tdict = {}
    for k, v in perf_info.items():
        col = 0
        if re.search("DL", k):
            for listinfo in perf_info["DL"]:
                n = listinfo.split("~")[1].strip()
                if n in tdict.keys():
                    tdict[n] = tdict[n] + 1
                elif n not in tdict.keys():
                    tdict[n] = 1

    row = row - 3

    workbook1 = xlsxwriter.Workbook("CPLX_Level_Monthly_Stats_" + str(formatted_xls_last_month) + ".xlsx")

    cell_format = workbook1.add_format({
        'bold': False,
        'font_color': 'black',
        'font_size': 11,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'text_wrap': True
    })

    table_header_format = workbook1.add_format({
        'bold': True,
        'font_color': 'black',
        'bg_color': 'red',
        'font_size': 11,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    topper_format = workbook1.add_format({
        'bold': True,
        'font_color': 'black',
        'bg_color': 'green',
        'font_size': 11,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    header_format = workbook1.add_format({
        'bold': True,
        'font_color': 'black',
        'bg_color': 'white',
        'font_size': 20,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    sub_header_format = workbook1.add_format({
        'bold': True,
        'font_color': 'black',
        'bg_color': 'white',
        'font_size': 14,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    worksheet1 = workbook1.add_worksheet()
    worksheet1.set_column("A:A", 20, cell_format)
    worksheet1.set_column("B:B", 14, cell_format)
    worksheet1.set_column("C:C", 14, cell_format)
    worksheet1.set_column("D:D", 14, cell_format)
    worksheet1.set_column("E:E", 14, cell_format)
    worksheet1.set_column("F:F", 35, cell_format)
    worksheet1.set_column("G:G", 35, cell_format)
    worksheet1.insert_image('A1', 'Fortinet-logomark-rgb-red.png',
                           {'x_scale': 0.75, 'y_scale': 0.75, 'x_offset': 30, 'y_offset': 10})
    worksheet1.insert_image('G1', 'Fortinet-logo-rgb-black-red.png',
                           {'x_scale': 1.25, 'y_scale': 1.25, 'x_offset': 5, 'y_offset': 15})
    merge_format = workbook1.add_format({'align': 'center', 'bold': True})

    worksheet1.merge_range('B2:F2', "Generated On - " + str(today) + "[" + str(creds['startdate']) + " - " + str(
        creds['enddate']) + "]", sub_header_format)
    worksheet1.merge_range('B1:F1', "Monthly Report : " + str(formatted_last_month), header_format)
    worksheet1.merge_range('B3:F3', "Queues Covered: AMER_SOAR, AMER_FMG_FAZ", sub_header_format)
    worksheet1.merge_range('A1:A3', "")
    worksheet1.merge_range('G1:G3', "")

    row = 3

    worksheet1.write(row, 0, 'NAME', table_header_format)
    worksheet1.write(row, 1, 'COUNT', table_header_format)
    worksheet1.merge_range(row, 2, row, 6, "DL TICKETS INFO", table_header_format)

    row = row + 1
    for k, v in tdict.items():
        col = 0
        worksheet1.write(row, col, k)
        worksheet1.write(row, col + 1, v)
        worksheet1.merge_range(row, col + 2, row, col +6, ' '.join(x.strip().split("~")[0] for x in perf_info['DL'] if re.search(k, x)))
        row = row + 1

    workbook1.close()

convert_to_excel("NULL")
exit()
def get_tkt_timer_info(url):
    print(url)
    driver.get(url)
    timer_tbl = driver.find_element(By.ID, "Table8")
    timer_tblrows = timer_tbl.find_elements(By.TAG_NAME, "tr")

    for row in timer_tblrows:
        cols = row.find_elements(By.TAG_NAME, "td")
        data = [column.text.strip() for column in cols]
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
    closeFromDate.send_keys(creds['startdate'])
    closeToDate.send_keys(creds['enddate'])
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