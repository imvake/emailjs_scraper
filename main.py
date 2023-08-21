from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import time
import pandas as pd
import re

driver = webdriver.Chrome()
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)


driver.get('https://dashboard.emailjs.com/sign-in')

time.sleep(5)


username = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="sing-in_email"]')))
username.clear()
username.click()
username.send_keys("")  # Your EmailJS email


pwd = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="sing-in_password"]')))
pwd.clear()
pwd.click()
pwd.send_keys("")  # Your EmailJS Password


submit_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="sing-in"]/button')))

submit_button.click()

# time.sleep(8)


email_history = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/section/div/section[1]/ul/li[4]/a/section')))
email_history.click()

time.sleep(5)
clickable_css_selector = '._row_bnx66_1'


WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, clickable_css_selector))
)

elements_with_class = driver.find_elements(
    By.CSS_SELECTOR, clickable_css_selector)


css_selector = '._section_1w0r6_1._description_1yq6b_6'


sections_list = []

records = []

for element in elements_with_class:
    element.click()
    time.sleep(1)
    sections = driver.find_elements(By.CSS_SELECTOR, '._section_1w0r6_1')
    for section in sections:
        text = section.text
        print("Section Text:", text)
        name_match = re.search(r'name:\s*(.*)', text)
        email_match = re.search(r'email:\s*(.*)', text)
        phone_match = re.search(r'phone:\s*(.*)', text)
        name = name_match.group(1).strip() if name_match else ''
        email = email_match.group(1).strip() if email_match else ''
        phone = phone_match.group(1).strip() if phone_match else ''
        records.append({
            'Name': name,
            'Email': email,
            'Phone': phone
        })
# *********** You can modify your req by adding values, (depends on your template) ***********
df = pd.DataFrame(records)


df = df.drop_duplicates()
df.to_excel('name.xlsx', index=False)


driver.quit()
