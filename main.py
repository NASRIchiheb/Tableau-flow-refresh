# Python refreshing Script
""""
"Tableau refreshing script based on Selenium and google sheets API (and some VBA macros)"

V0.0001 pre-Alpha
-------------------------------------
Selenium : Chrome driver / Logs in code
Google sheets : Token in Keys.json file
-------------------------------------
Execution process : 

Requirements :
 - Python 3+
 - Selenium
 - Google api client
 - Keys.json file, for api connection
 - !!!!!!!!!!!!!!!!! Tableau online should be in french !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

How it works :
Launch main.py

1- Get info from google sheets :
    - Flow url, update schedule , status
    - Launch selenium , Execute flow, get status
    - Wait for flow execution.
        -If success goes to next line and update status / if fail goes to next line and send email

The management of the flows is based on google sheet : 
https://docs.google.com/spreadsheets/d/1cMaet7jXVZ_7eLYgr7ZQp5CyWG-GtEsVHHZUGVwCXg0/edit#gid=0

A vba script is also used to send automatic emails in case of fails.

----------------------------------------
V0.0001 Features:
- Read Drive sheet.
- Execute "Daily flows"
- Execute entire flow (not one or 2 outputs)
- Keep status of refreshing inside google sheet
- Send automatic e-mail in case of fail
----------------------------------------

Improvements :
 - If fail start again 2 times :  DONE
 - Optimize waiting times
 - Test with long executing flows
 - Launch only 1 output and not execute all the flow outputs :  DONE
 - Test with a full list of flows : DONE
 - Complete error emails with more details
 - Handle more update schedules (Currently only daily flows)
 - Automate flow execution ( Task manager ?)
 - Make it usable in a server
 - Add time of refreshing : DONE
"""

# Import packages
import time
from datetime import datetime

# Google sheets packages
from googleapiclient.discovery import build
from google.oauth2 import service_account

# Packages for selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException

# Connect to google sheets to access the links

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# Keys.json contains connection tokens and info of sheet API
# The file keys.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.

SERVICE_ACCOUNT_FILE = 'keys.json'

# Initialise credentials
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID of a sample spreadsheet. Contained in the sheet URL
spreadsheet_id = '1cMaet7jXVZ_7eLYgr7ZQp5CyWG-GtEsVHHZUGVwCXg0'

# Make the connection
service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()

# Get the the first column of sheet containing the links
result = sheet.values().get(spreadsheetId=spreadsheet_id,
                            range="Testsheet1!A:K").execute()

# Update job status
now = datetime.now()
jobstart = time.time()
ao = [['Job Start', now.strftime("%d/%m/%Y %H:%M:%S")]]
sheet.values().update(spreadsheetId=spreadsheet_id, range="Testsheet1!N2",
                      valueInputOption="USER_ENTERED", body={"values": ao}).execute()

# flows contains list of flows starting from index 1
flows = result.get('values', [])
qtyOfFlows = len(flows) - 1

# ----------------------------------------
# Connect to tableau with selenium and start updating

# Open chrome with selenium
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

# Loop through all the list of flows and update them one by
# Remove first element of the list because it's the header of the column
flows.pop(0)


# functions -------------------------------------------------------------------------

def executeall(indexnumber):
    # Click on execute all
    driver.find_element(By.CSS_SELECTOR, "[title*='Tout exécuter']").click()
    time.sleep(1)

    # Confirm execute
    driver.find_element(By.XPATH,
                        '//*[@id="react-confirm-action-dialog-Dialog-Body-Id"]/div[2]/span/button').click()

    # Update google sheet status
    start_exec = datetime.now()
    ao = [['Flow executing', start_exec.strftime("%d/%m/%Y %H:%M:%S")]]

    sheet.values().update(spreadsheetId=spreadsheet_id, range="Testsheet1!G" + str(indexnumber + 2),
                          valueInputOption="USER_ENTERED", body={"values": ao}).execute()
    time.sleep(1)
    driver.refresh()
    time.sleep(2)

    print("waiting for pending")
    # Wait until execution is done to get the status

    while len(driver.find_elements(By.CLASS_NAME, "tb-prep-flow-overview-history-status-pending-icon")) != 0:
        driver.refresh()
        time.sleep(30)

    print("Fin while")

    time.sleep(1)
    WebDriverWait(driver, 3600).until_not(
        lambda x: x.find_element(By.CLASS_NAME, "tb-prep-flow-overview-history-status-pending-icon").is_displayed())

    while len(driver.find_elements(By.CLASS_NAME, "tb-prep-flow-overview-history-status-inprogress-icon")) != 0:
        driver.refresh()
        time.sleep(30)

    print("Fin while")

    time.sleep(1)

    while len(driver.find_elements(By.CLASS_NAME, "tb-prep-flow-overview-history-status-inprogress-icon")) != 0:
        driver.refresh()
        time.sleep(30)

    time.sleep(10)
    driver.refresh()
    time.sleep(5)
    WebDriverWait(driver, 3600).until_not(
        lambda x: x.find_elements(By.CLASS_NAME, "tb-prep-flow-overview-history-status-inprogress-icon"))

    print("End of pending")

    # Refresh page to be sure flow status appears

    driver.refresh()
    time.sleep(1)


def executeoutput(indexnumber, output):
    time.sleep(1)

    # Get all execute buttons and choose output by its index
    elements = driver.find_elements(By.XPATH, '//button[@class="f1odzkbq f1p7b3nz low-density"]')
    elements[int(output)].click()

    # Confirm execute
    driver.find_element(By.XPATH,
                        '//*[@id="react-confirm-action-dialog-Dialog-Body-Id"]/div[2]/span/button').click()

    now = datetime.now()
    ao = [['Flow executing', now.strftime("%d/%m/%Y %H:%M:%S")]]

    sheet.values().update(spreadsheetId=spreadsheet_id, range="Testsheet1!G" + str(indexnumber + 2),
                          valueInputOption="USER_ENTERED", body={"values": ao}).execute()
    time.sleep(1)
    driver.refresh()
    time.sleep(2)

    print("waiting for pending")
    # Wait until execution is done to get the status

    while len(driver.find_elements(By.CLASS_NAME, "tb-prep-flow-overview-history-status-pending-icon")) != 0:
        driver.refresh()
        time.sleep(15)

    print("Fin while")

    WebDriverWait(driver, 3600).until_not(
        lambda x: x.find_elements(By.CLASS_NAME, "tb-prep-flow-overview-history-status-pending-icon"))

    print("End of pending")
    time.sleep(10)
    driver.refresh()
    time.sleep(5)

    while len(driver.find_elements(By.CLASS_NAME, "tb-prep-flow-overview-history-status-inprogress-icon")) != 0:
        driver.refresh()
        time.sleep(30)

    print("Fin while")


    WebDriverWait(driver, 3600).until_not(
        lambda x: x.find_elements(By.CLASS_NAME, "tb-prep-flow-overview-history-status-inprogress-icon"))

    # Refresh page to be sure flow status appears

    driver.refresh()


def runallflow(indexnumber, flowslist):
    start_time = 0
    end_time = 0
    # Go to your page
    driver.get(flowslist[indexnumber][0])
    time.sleep(1)
    try:
        # Log in
        input_element = driver.find_element(By.ID, "email")
        input_element.send_keys('andreisharfenberg@swap-europe.com')
        input_element.send_keys(Keys.ENTER)
        time.sleep(1)
        input_element = driver.find_element(By.ID, "password")
        input_element.send_keys('aB12345!')
        input_element.send_keys(Keys.ENTER)
        time.sleep(5)
        # Close update popup
        driver.find_element(By.XPATH, '//*[@id="app-root"]/div[3]/div/div/div/div/div[2]/div/div/button').click()
    except NoSuchElementException:
        pass

    time.sleep(1)
    start_time = time.time()
    exec_start = datetime.now()
    executeall(indexnumber)

    time.sleep(1)

    # Try the different status possibilities : Fail, Success
    # To add cancelled flow status
    # Use try catch to avoid errors
    element = None
    for _ in range(1):
        try:
            # If not success , it's fail (or cancelled to be added)
            element = WebDriverWait(driver, 10).until(lambda x: x.find_element(By.CLASS_NAME,
                                                                               "tb-prep-flow-overview-history"
                                                                               "-status-success-icon").is_displayed())
            now = datetime.now()
            end_time = time.time()
            # In case of success update sheet with status and time
            print_result = [['Success', exec_start.strftime("%d/%m/%Y %H:%M:%S"), now.strftime("%d/%m/%Y %H:%M:%S"),
                             ((end_time - start_time) / 60) - 1, now.strftime("%d/%m/%Y %H:%M:%S")]]

            sheet.values().update(spreadsheetId=spreadsheet_id, range="Testsheet1!G" + str(indexnumber + 2),
                                  valueInputOption="USER_ENTERED", body={"values": print_result}).execute()
            break
        except (TimeoutException, NoSuchElementException):
            print("flow failed" + flowslist[indexnumber][1])
            executeall(indexnumber)

    else:
        pass

    try:
        # If not success , it's fail (or cancelled to be added)
        element = WebDriverWait(driver, 10).until(lambda x: x.find_element(By.CLASS_NAME,
                                                                           "tb-prep-flow-overview-history"
                                                                           "-status-failed-icon").is_displayed())
        # In case of success update sheet with status and time
        print_result = [['Fail', "", ""]]
        sheet.values().update(spreadsheetId=spreadsheet_id, range="Testsheet1!G" + str(indexnumber + 2),
                              valueInputOption="USER_ENTERED", body={"values": print_result}).execute()
    except (TimeoutException, NoSuchElementException):
        pass

    return print("End of execution for flow")


def runoutpuflow(indexnumber, flowslist, output):
    # Go to your page
    driver.get(flowslist[indexnumber][0])
    time.sleep(1)
    try:
        # Log in
        input_element = driver.find_element(By.ID, "email")
        input_element.send_keys('andreisharfenberg@swap-europe.com')
        input_element.send_keys(Keys.ENTER)
        time.sleep(1)
        input_element = driver.find_element(By.ID, "password")
        input_element.send_keys('aB12345!')
        input_element.send_keys(Keys.ENTER)
        time.sleep(5)
        # Close update popup
        driver.find_element(By.XPATH, '//*[@id="app-root"]/div[3]/div/div/div/div/div[2]/div/div/button').click()
    except NoSuchElementException:
        pass

    start_time = time.time()
    exec_start = datetime.now()

    time.sleep(1)

    executeoutput(indexnumber, output)

    time.sleep(5)

    element = None

    # Try the different status possibilities : Fail, Success
    # To add cancelled flow status
    # Use try catch to avoid errors

    for _ in range(1):
        try:
            # If not success , it's fail (or cancelled to be added)
            element = WebDriverWait(driver, 10).until(lambda x: x.find_element(By.CLASS_NAME,
                                                                               "tb-prep-flow-overview-history"
                                                                               "-status-success-icon").is_displayed())
            now = datetime.now()
            end_time = time.time()
            # In case of success update sheet with status and time
            print_result = [['Success', exec_start.strftime("%d/%m/%Y %H:%M:%S"), now.strftime("%d/%m/%Y %H:%M:%S"),
                             ((end_time - start_time) / 60) - 1, now.strftime("%d/%m/%Y %H:%M:%S")]]
            sheet.values().update(spreadsheetId=spreadsheet_id, range="Testsheet1!G" + str(indexnumber + 2),
                                  valueInputOption="USER_ENTERED", body={"values": print_result}).execute()
            break
        except (TimeoutException, NoSuchElementException):
            executeoutput(indexnumber, output)
    else:
        pass

    try:
        # If not success , it's fail (or cancelled to be added)
        element = WebDriverWait(driver, 10).until(lambda x: x.find_element(By.CLASS_NAME,
                                                                           "tb-prep-flow-overview-history"
                                                                           "-status-failed-icon").is_displayed())
        # In case of success update sheet with status and time
        print_result = [['Fail', "", ""]]
        sheet.values().update(spreadsheetId=spreadsheet_id, range="Testsheet1!G" + str(indexnumber + 2),
                              valueInputOption="USER_ENTERED", body={"values": print_result}).execute()
    except (TimeoutException, NoSuchElementException):
        pass

    return print("End of execution for flow")


def day_of_week():
    today_date = datetime.today().weekday()
    if today_date == 5:
        return "weekend"
    elif 0 <= today_date < 5:
        return "weekday"


"""
Loop execution start:
"""

for i in range(qtyOfFlows):  # Loop through the list of flows
    time.sleep(5)
    # Daily refreshing
    if flows[i][5] == "Daily (working days)" and day_of_week() == "weekday":  # If it's daily flow so we execute
        if flows[i][2] == "Yes":
            runallflow(i, flows)
        elif flows[i][2] == "No":
            print("output exe")
            runoutpuflow(i, flows, flows[i][3])

    # Weekly
    elif flows[i][5] == "Weekly (on weekend)" and day_of_week() == "weekend":
        if flows[i][2] == "Yes":
            runallflow(i, flows)
        elif flows[i][2] == "No":
            print("output exe")
            runoutpuflow(i, flows, flows[i][3])

now = datetime.now()
jobend = time.time()
# In case of success update sheet with status and time
ao = [['Job complete', now.strftime("%d/%m/%Y %H:%M:%S"), "Total execution time", ((jobend - jobstart) / 60) - qtyOfFlows]]
sheet.values().update(spreadsheetId=spreadsheet_id, range="Testsheet1!P2",
                      valueInputOption="USER_ENTERED", body={"values": ao}).execute()

driver.close()
print("END OF SCRIPT")
