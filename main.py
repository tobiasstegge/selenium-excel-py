from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions

import sys
import pandas as pd


url = sys.argv[1]

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--incognito")
options.add_argument("window-size=1201,1400")

caps = DesiredCapabilities().CHROME
caps["pageLoadStrategy"] = "eager"

driver = webdriver.Chrome('./drivers/mac/chromedriver', options=options, desired_capabilities=caps)
page = driver.get(url=url)

WebDriverWait(driver, 5).until(expected_conditions.element_to_be_clickable((
    By.CLASS_NAME, 'i-result-item')))

driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/button[1]').click()
driver.implicitly_wait(3)
results = driver.find_elements_by_class_name('i-result-item')

df = pd.DataFrame(columns=['name', 'homepage', 'phone', 'info', 'contact_person'])

for result in results:

    driver.implicitly_wait(1)

    WebDriverWait(driver, 5).until(expected_conditions.element_to_be_clickable((
        By.CLASS_NAME, 'i-result-item')))

    result.click()
    WebDriverWait(driver, 5).until(expected_conditions.element_to_be_clickable((
        By.CLASS_NAME, 'ui-button-text')))
    result.find_element_by_class_name('ui-button-text').click()

    driver.implicitly_wait(0.5)

    WebDriverWait(driver, 10).until(expected_conditions.element_to_be_clickable((
        By.CLASS_NAME, 'i-pagetitle')))

    # get all elements from company

    name = driver.find_element_by_class_name('i-pagetitle').text
    homepage = driver.find_element_by_class_name('i-link-external').get_attribute('href')
    phone = driver.find_element_by_xpath('//*[@id="profile"]/div[4]/div[2]/div/div/div/div[4]').text
    info = driver.find_element_by_class_name('ui-accordion-content').text
    driver.find_element_by_class_name('i-accordion-anbieter').click()
    WebDriverWait(driver, 5).until(expected_conditions.element_to_be_clickable((
        By.CLASS_NAME, 'i-ansprechpartner')))
    contact_person = driver.find_element_by_class_name('i-ansprechpartner').text



    # Create a Pandas dataframe from some data.
    df = df.append({
        'name': name[4:],
        'homepage': homepage,
        'phone': phone,
        'info': info,
        'contact_person': contact_person
    }, ignore_index=True)
    driver.back()

    print(df)


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer.save()
