from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import pandas as pd
import os

web = "https://app.clientify.com/accounts/login/?next=/"
path = "E:\STUDY\STUDY\PYTHON\Web Scraping\ClientifyReport\chromedriver.exe"
service = Service(executable_path=path)
driver = webdriver.Chrome(service=service)
driver.get(web)
driver.maximize_window()

# Using the username textbox
username_input = driver.find_element(By.XPATH,
                                     '//input[@name="username"]')
username_input.send_keys("sistemas@mantaro.pe")

# Using the password textbox
password = driver.find_element(By.XPATH,
                               '//input[@name="password"]')
password.send_keys("k4gHOSwU*")

# Using the login button
login_button = driver.find_element(By.XPATH,
                                   '//button[@type="submit"]')
login_button.click()
time.sleep(4)

contact_button = driver.find_element(By.XPATH,
                                     '//a[@href="/contacts/"]')
contact_button.click()
time.sleep(2)

order_by_creation_button = driver.find_element(By.XPATH,
                                               '//a[@data-value="created"]')
order_by_creation_button.click()
time.sleep(2)

order_by_creation_button = driver.find_element(By.XPATH,
                                               '(//a[@data-value="-created"])[2]')
order_by_creation_button.click()
time.sleep(2)

owners = []
states = []
leads = []


def process():
    next_button = driver.find_element(By.XPATH, './/ul[@class="pagination "]//li[8]//a')
    rows = driver.find_elements(By.XPATH, '//tr[@class="list-edit-row"]')
    for row in rows:
        time_creation = row.find_element(By.XPATH, '(.//td[@class="hidden-xs"])[4]//div').text
        print(time_creation)
        if time_creation.__contains__("semana"):
            return

        owner = row.find_element(By.XPATH, '(.//td[@class="hidden-xs"])[3]').text
        status = row.find_element(By.XPATH, '(.//td[@class="hidden-xs"])[2]//span').text
        lead = row.find_element(By.XPATH, '(.//a[@class="list-full-name"])[1]').get_attribute("title")

        owners.append(owner)
        states.append(status)
        leads.append(lead)

    next_button.click()
    time.sleep(2)
    process()

process()

driver.quit()
df_data = pd.DataFrame({"owners": owners, "states": states, "leads": leads})
df_data.to_excel("data.xlsx")
print(df_data)
