from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
from datetime import datetime, timedelta

import os

service = Service(executable_path=r"C:\Users\kushal\Downloads\chromedriver\chromedriver.exe")
driver = webdriver.Chrome(service=service)

driver.get("https://cloud.logicerpcloud.com/Login/UserLogin")
time.sleep(2)
username = driver.find_element(By.XPATH, "/html/body/div/form/div[2]/div[1]/input")
passwd = driver.find_element(By.XPATH, "/html/body/div/form/div[2]/div[2]/input")

username.send_keys("DEVAANX_1")
passwd.send_keys("Admin@123")
time.sleep(1)
########################## First Login ###########################
driver.find_element(By.XPATH, "/html/body/div/form/div[2]/div[5]/div/button").click()
time.sleep(1)
driver.find_element(By.XPATH,"/html/body/div[3]/div[2]/div/div[1]/div/div/div").click()
time.sleep(1)
############################### Second Login #######################################
driver.switch_to.window(driver.window_handles[-1])
new_url = driver.current_url
# driver.get(f"{new_url}")
time.sleep(1)
# driver.switch_to.new_window('tab')
driver.find_element(By.XPATH,"/html/body/div[3]/fieldset[2]/div[2]/div/div/div/div[1]").click() # Selct Dropdown
time.sleep(1)
driver.find_element(By.XPATH,"/html/body/div[6]/div/div/div/div[2]/div/div[2]").click() #Select User
time.sleep(1)
passwd = driver.find_element(By.XPATH,"/html/body/div[3]/fieldset[2]/div[3]/input")
passwd.send_keys("Admin@123")
driver.find_element(By.XPATH,"/html/body/div[3]/fieldset[2]/div[4]/input[2]").click() # Click Log In
time.sleep(1)
driver.find_element(By.XPATH,"/html/body/div[3]/fieldset[3]/div[6]/input[2]").click() # Click ok
time.sleep(2)
driver.find_element(By.XPATH,"/html/body/div[3]/fieldset[4]/div[6]/div/ul/li[3]/div").click() # Select devaa nx
time.sleep(1)
driver.find_element(By.XPATH,"/html/body/div[3]/fieldset[4]/div[7]/input[2]").click() # Click ok
# ######################### Now we are LOGGED IN #####################

# # Select My Menu
time.sleep(5)
driver.find_element(By.XPATH,"/html/body/div[3]/div[3]/div[1]/ul/li[2]").click() # Click My Menu
time.sleep(1)
# Select Reports
driver.find_element(By.XPATH,"/html/body/div[3]/div[3]/div[2]/div[2]/table/tbody/tr[1]/td/ul/li[2]").click() # Click Reports
time.sleep(1)
# Select Devaa Annex Reports
driver.find_element(By.XPATH,"/html/body/div[3]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/div/div[2]/div/nav/div/ul/li[10]").click() # Click Devaa Annex Reports
time.sleep(1)
# Select "Detailed Sales Register"
driver.find_element(By.XPATH,"/html/body/div[3]/div[3]/div[2]/div[2]/table/tbody/tr[2]/td/div/div[2]/div/nav/div/ul/li[10]/div/ul/li[2]").click() # Click Detailed Sales Register
time.sleep(1)
# Switch to another Tab
driver.switch_to.window(driver.window_handles[-1])
time.sleep(1)
# Select Start Date
date_start = driver.find_element(By.XPATH,"/html/body/div[3]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[2]/td/table/tbody/tr[6]/td/div/table/tbody/tr/td[2]/div/div/input") # Select Start Date
date_start.click()
time.sleep(1)
day = (datetime.now() - timedelta(1)).strftime("%d")
month = (datetime.now() - timedelta(1)).strftime("%m")
year = (datetime.now() - timedelta(1)).strftime("%Y")
date_start.send_keys("25")
date_start.send_keys("04")
date_start.send_keys(year)
time.sleep(1)

# Select End Date
date_end = driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[2]/td/table/tbody/tr[6]/td/div/table/tbody/tr/td[4]/div/div/input") #Select End date
date_end.click()
time.sleep(1)
date_end.send_keys("27")
date_end.send_keys("04")
date_end.send_keys(year)

# Select Filter
driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/div[1]/div/div[1]/button").click() # Click Filter button
time.sleep(1)

# Select Branch
driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/div[1]/div/div[1]/table/tbody/tr[1]/td/div/div/div[2]/div/div[3]/div[2]/div/div[4]/div[1]/div").click() # Click Branch filter
time.sleep(1)

# Select Devaa Annex
driver.find_element(By.XPATH, "/html/body/div[3]/div[5]/div/div/div[2]/div[2]/div/table/tbody/tr[5]/td/div/div[2]/div/div[3]/div[2]/div/div[3]/div[1]/div").click() # Click Devaa Annex as a branch
time.sleep(1)

# Click ok
driver.find_element(By.XPATH, "/html/body/div[3]/div[5]/div/div/div[2]/div[2]/div/table/tbody/tr[6]/td[6]").click() # Click ok
time.sleep(1)

# Create the report
driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/button[1]").click() # Click to create the report
time.sleep(1)

# Wait until page navigation isn't there /html/body/div[3]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr[2]/td/nav/ul
# WebDriverWait(driver=driver, timeout=20).until(EC.visibility_of((By.XPATH, "/html/body/div[3]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr[2]/td/nav/ul")))
time.sleep(30)
# Export the report into excel
driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr[1]/td/div/div[2]/div/div[1]/div/input[1]").click() # Click to export the report as Excel

# Select Branch Filter again
driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/div[1]/div/div[1]/table/tbody/tr[1]/td/div/div/div[2]/div/div[3]/div[2]/div/div[4]/div[1]/div").click() # Click Branch filter
time.sleep(2)
driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/div[1]/div/div[1]/table/tbody/tr[1]/td/div/div/div[2]/div/div[3]/div[2]/div/div[4]/div[1]/div").click() # Click Branch filter
time.sleep(1)

# Deselect Devaa Annex
driver.find_element(By.XPATH, "/html/body/div[3]/div[6]/div/div/div[2]/div[2]/div/table/tbody/tr[5]/td/div/div[2]/div/div[3]/div[2]/div/div[3]/div[1]/div").click() # Click Devaa Annex to deselect
time.sleep(1)

# Select Devaa for Women
driver.find_element(By.XPATH, "/html/body/div[3]/div[6]/div/div/div[2]/div[2]/div/table/tbody/tr[5]/td/div/div[2]/div/div[3]/div[2]/div/div[4]/div[1]/div").click() # Click Devaa for women as a branch
time.sleep(1)

# Click ok
driver.find_element(By.XPATH, "/html/body/div[3]/div[6]/div/div/div[2]/div[2]/div/table/tbody/tr[6]/td[6]/input").click() # Click ok
time.sleep(1)

# Create the report
driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/button[1]").click() # Click to create the report
time.sleep(30)

# Export the report into excel
driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr[1]/td/div/div[2]/div/div[1]/div/input[1]").click() # Click to export the report as Excel
time.sleep(5)
driver.quit()



