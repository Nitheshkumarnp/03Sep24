#   This file will run the automation of downloading pdf from Gemini live.

import logging
import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl

# Configuring logs
logging.basicConfig(
    filename='logs.log',    # Log file location
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO      # Set the logging level
)

# Read data from excel file and store it in list
def read_ids_from_excel(file_path, lower_limit, upper_limit):
    '''
    This method will read data from excel file. Based on the lower and upper limit, data from excel will be added to ids list and ids will be returned as output of the method.

    Parameters:

        file_path = path of the selected file,
        lower_limit = minimum row of the excel file,
        upper_limit = maximum row of the excel file
    
    '''
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    ids = []
    for row in sheet.iter_rows(min_row=lower_limit, max_row=upper_limit, max_col=1, values_only=True):  # Assuming IDs are in the first column
        ids.append(row[0])
    return ids

# This is used to switch the mode from debug and development automatically based on system arguments.
debug_mode = True if len(sys.argv) < 2 else False

# required inputs
if debug_mode:
    # Hardcoding data required to test manually without starting the application
    excel_file_path = "PNR_list.xlsx"
    lower_limit = 2
    upper_limit = 60
    timeToLoad = 150
else:
    # Original data comes from api request
    excel_file_path = sys.argv[1]
    lower_limit = int(sys.argv[2])
    upper_limit = int(sys.argv[3])
    timeToLoad = int(sys.argv[4])

# Gets list of id values come from method's output
ids_to_search = read_ids_from_excel(excel_file_path, lower_limit, upper_limit)

# Adding id list to logs and removing duplicates
logging.info(f'Total Number of ids = {len(ids_to_search)}')
logging.info(ids_to_search)
ids_to_search = list(dict.fromkeys(ids_to_search))
logging.info(f'Unique number of ids = {len(ids_to_search)}')
logging.info(ids_to_search)

# Initialize the WebDriver
driver = webdriver.Chrome()

# URL of the web page
url = "https://fcmindia.okta.com/"

# Maximize the screen size
driver.maximize_window()

# Open the web page
driver.get(url)

# Mail id
WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-username"]')))
mailId = driver.find_element(By.XPATH, '//*[@id="okta-signin-username"]')
mailId.send_keys('ganesh.ss@fcmin.com')
mailId.send_keys(Keys.TAB)

# Mail id password
WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-password"]')))
mailPassword = driver.find_element(By.XPATH, '//*[@id="okta-signin-password"]')
mailPassword.send_keys('Ganu@87654321')
mailPassword.send_keys(Keys.RETURN)

# Push Button
WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="form63"]/div[2]/input')))
pushButton = driver.find_element(By.XPATH, '//*[@id="form63"]/div[2]/input')
pushButton.click()

# PHX Booking India
WebDriverWait(driver, 300).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/section/section/section/div/section/div[14]')))
fcmTab = driver.find_element(By.XPATH, '//*[@id="main-content"]/section/section/section/div/section/div[14]')
fcmTab.click()

# Waiting time to close the login tab
time.sleep(20)

# Get handles of all currently open windows
window_handles = driver.window_handles

# Switch to the new tab (which is the latest one opened)
new_tab_handle = window_handles[-1]
driver.switch_to.window(new_tab_handle)

# Selecting Profile Name
WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="menuProfileForm:roleMenu"]/ul')))
profileName = driver.find_element(By.XPATH, '//*[@id="menuProfileForm:roleMenu"]/ul')

# Create an instance of ActionChains
action = ActionChains(driver)

# Perform the hover action
action.move_to_element(profileName).perform()

# Selecting V - Vendor
WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="menuProfileForm:roleMenu"]/ul/li/ul/li[3]')))
vendor = driver.find_element(By.XPATH, '//*[@id="menuProfileForm:roleMenu"]/ul/li/ul/li[3]')
vendor.click()

# Count for total number of invoice downloaded
totalInvoiceDownloaded = 0

# Storing Cart Number not found
CartNumbersNotFound = []

# Common automation Script for all the 4 tabs
def automation(idList):
    global totalInvoiceDownloaded
    logging.info("Entering Automation method")

    # This try block is implemented to avoid the time and data loss, while uncertain stop in automation function
    try:

        # Iterating PNR numbers
        for index, search_id in enumerate(idList):
            if search_id == None:
                break

            # Click on Search By
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="travelDashboardForm:searchFilter_label"]')))
            time.sleep(1)
            searchBy = driver.find_element(By.XPATH, '//*[@id="travelDashboardForm:searchFilter_label"]')
            searchBy.click()

            # Selecting Airline PNR
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="travelDashboardForm:searchFilter_2"]')))
            time.sleep(1)
            airlinePNR = driver.find_element(By.XPATH, '//*[@id="travelDashboardForm:searchFilter_2"]')
            airlinePNR.click()

            # Click on Search Value
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="travelDashboardForm:searchInputText"]')))
            time.sleep(1)
            searchValue = driver.find_element(By.XPATH, '//*[@id="travelDashboardForm:searchInputText"]')
            searchValue.send_keys(search_id)

            # Click on Search Button
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/div[2]/div/div[3]/div/form/div[3]/div[4]/div/button')))
            searchButton = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/div/div[3]/div/form/div[3]/div[4]/div/button')
            searchButton.click()

            # Checking for Cart Number
            try:

                # Click on Emulate Button
                WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/div[2]/div/div[3]/div/form/div[6]/div/table/tbody/tr/td[7]/button')))
                emulateButton = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/div/div[3]/div/form/div[6]/div/table/tbody/tr/td[7]/button')
                emulateButton.click()

                # Click on Show Detail
                WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="travelDashboardForm:tripListDT_data"]/tr/td[7]/div[2]')))
                showDetail = driver.find_element(By.XPATH, '//*[@id="travelDashboardForm:tripListDT_data"]/tr/td[7]/div[2]')
                showDetail.click()

                # Click on View Summary
                WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="travelDashboardForm:tripListDT:0:viewSummaryID"]')))
                viewSummary = driver.find_element(By.XPATH, '//*[@id="travelDashboardForm:tripListDT:0:viewSummaryID"]')
                viewSummary.click()

                # Checking for Invoice Number
                try:

                    # Counting number of invoices
                    WebDriverWait(driver, 15).until(EC.visibility_of_all_elements_located((By.XPATH, '/html/body/div[1]/div[1]/div[2]/form/div[1]/div[3]/div[1]/div[4]/div/div/div/div[2]/div/div[2]/div/div[1]/table/tbody/tr')))
                    totalNumberOfRows = driver.find_elements(By.XPATH, '/html/body/div[1]/div[1]/div[2]/form/div[1]/div[3]/div[1]/div[4]/div/div/div/div[2]/div/div[2]/div/div[1]/table/tbody/tr')
                    
                    # Iterating to download invoices
                    for row in range(len(totalNumberOfRows)):
                        # Downloading Invoice
                        WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, f'//*[@id="tripSummary:serviceTabView:airInvoiceList:{row}:ticketDownloadFile"]')))
                        clientInvoice = driver.find_element(By.XPATH, f'//*[@id="tripSummary:serviceTabView:airInvoiceList:{row}:ticketDownloadFile"]')
                        clientInvoice.click()
                        totalInvoiceDownloaded += 1
                        logging.info(f'Rows = {index+1}, PNR = {search_id}, Invoice count = {row+1}, Total Invoice Downloaded = {totalInvoiceDownloaded}')
                        time.sleep(2)
                
                except:
                    
                    # Logging for Invoice Number
                    logging.info(f'Rows = {index+1}, PNR = {search_id}, Invoice Number not found')
            
            except:
                
                # Adding Cart Number to list
                CartNumbersNotFound.append(search_id)

                # Logging for Cart Number
                logging.info(f'Rows = {index+1}, PNR = {search_id}, Cart Number not found')

            # Getting back to Main Screen
            driver.back()

            # Waiting to load Main Screen
            time.sleep(3)
        
        # Number of Invoice not found:
        logging.info(f'Number of Invoice not found :  {len(CartNumbersNotFound)}')

        # Cart Numbers not found List
        logging.info(f'Cart Numbers not found :  {CartNumbersNotFound}')

        # Returning the Updated result
        return f'Total Rows = {index+1}, Last PNR = {search_id}, Total Invoice downloaded = {totalInvoiceDownloaded}'

    except:

        # Number of Invoice not found:
        logging.info(f'Number of Invoice not found :  {len(CartNumbersNotFound)}')

        # Cart Numbers not found List
        logging.info(f'Cart Numbers not found :  {CartNumbersNotFound}')
        
        # Returning the Updated result
        return f'Total Rows = {index+1}, Last PNR = {search_id}, Total Invoice downloaded = {totalInvoiceDownloaded}'

# Running Automation
result = automation(ids_to_search)

# Print Result
print(result)