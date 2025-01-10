from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd

driver = webdriver.Chrome()

df = pd.read_excel('challenge.xlsx')
df.columns = df.columns.str.strip()
print(df.columns)

driver.get("https://www.rpachallenge.com/")

start_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Start')]")
start_button.click()

for index, row in df.iterrows():
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelAddress']").send_keys(row['Address'])
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelFirstName']").send_keys(row['First Name'])
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelLastName']").send_keys(row['Last Name'])
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelCompanyName']").send_keys(row['Company Name'])
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelRole']").send_keys(row['Role in Company'])
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelEmail']").send_keys(row['Email'])
    driver.find_element(By.XPATH, "//input[@ng-reflect-name='labelPhone']").send_keys(row['Phone Number'])

    submit_button = driver.find_element(By.XPATH, "//input[@type='submit']")
    submit_button.click()

time.sleep(20)
driver.quit()
