'''

The purpose of this is to:
1) take values out of a csv file,
2) plug them into and online survey
link: https://www.ressourcen-rechner.de/calculator.php?lang=en
3) scrape the result
4) insert back into python

notes:
The waiting and next button commands, while inelegant,
are necessary to execute the javascript survey

'''
#first install selenium, pandas, and xlrd
import csv
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.webelement import WebElement
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd

#reads file
dataframe = pd.read_csv('test_footprint1.csv')

'''
DATA MANIPULATION SECTION
To make the data fit into the online survey
some of it must be manipulated (particularly
for the radio buttons)
'''

dataframe['HO03'] = dataframe['HO03'].replace(['1','2','3','4'], [41,42,43,44])
dataframe['HO04'] = dataframe['HO04'].replace(['1','2','3','4'], [46,47,48,49])
dataframe['HO05'] = dataframe['HO05'].replace(['1','2','3','4','5','6','7','8'], [51,52,53,54,55,56,57,58])
dataframe['HO06'] = dataframe['HO06'].replace(['1','2','3','4'], [60,61,62,63])

# Testing start point
entry = dataframe.iloc[1]

# BROWSER SET UP
driver_path = 'C:/Users/Beemo/chromedriver/chromedriver.exe'
brave_path = 'C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe'

option = webdriver.ChromeOptions()
option.binary_location = brave_path

browser = webdriver.Chrome(executable_path=driver_path, options=option)
browser.get('https://www.ressourcen-rechner.de/calculator.php?lang=en/')

'''
FILLING THE SURVEY
'''

#pause for debugging
time.sleep(1)
                                        
#select language
lang_select = Select(browser.find_element_by_id('lang-selector'))
lang_select.select_by_value(entry['lang'])

time.sleep(1)

# question: place of residence
country_select = Select(browser.find_element_by_id('answers01'))
country_select.select_by_index(entry['HO01'])


time.sleep(1)
browser.find_element_by_css_selector('p[id="next-button"]').click()
time.sleep(1)

# Question 2 (input): # Adults, # children, & size of Household
adult_input = browser.find_element_by_id('37')
adult_input.send_keys(entry['HO02_01'])

child_input = browser.find_element_by_id('38')
child_input.send_keys(entry['HO02_02'])

house_size_input = browser.find_element_by_id('39')
house_size_input.send_keys(entry['HO02_03'])

time.sleep(1)
browser.find_element_by_css_selector('p[id="next-button"]').click()
time.sleep(1)

# Question 3 (radio): Electricity type
elec_type_xpath = f"//input[@name=\"Q03\"][@id=\"{entry['HO03']}\"]"
elec_type_radio = browser.find_element_by_xpath(elec_type_xpath)
browser.execute_script("return arguments[0].click()", elec_type_radio)

time.sleep(1)
browser.find_element_by_css_selector('p[id="next-button"]').click()
time.sleep(1)

# Question 4 (radio + input): Electricity usage
elec_usage_xpath = f"//input[@name=\"Q04\"][@id=\"{entry['HO04']}\"]"
elec_usage_radio = browser.find_element_by_xpath(elec_usage_xpath)
browser.execute_script("return arguments[0].click()", elec_usage_radio)

time.sleep(1)

if len(dataframe.HO04_01) > 0:
    elec_usage_input = browser.find_element_by_name('optValQ04')
    elec_usage_input.send_keys(entry['HO04_01'])

time.sleep(1)
browser.find_element_by_css_selector('p[id="next-button"]').click()
time.sleep(1)

# Question 5 (radio): Heating type
heat_type_xpath = f"//input[@name=\"Q05\"][@id=\"{entry['HO05']}\"]"
heat_type_radio = browser.find_element_by_xpath(heat_type_xpath)
browser.execute_script("return arguments[0].click()", heat_type_radio)

time.sleep(1)
browser.find_element_by_css_selector('p[id="next-button"]').click()
time.sleep(1)

# Question 6 (radio + input): Heating usage
heat_usage_xpath = f"//input[@name=\"Q06\"][@id=\"{entry['HO06']}\"]"
heat_usage_radio = browser.find_element_by_xpath(heat_usage_xpath)
browser.execute_script("return arguments[0].click()", heat_usage_radio)

time.sleep(1)

if len(dataframe.HO06_01) > 0:
    heat_usage_input = browser.find_element_by_name('optValQ06')
    heat_usage_input.send_keys(entry['HO06_01'])

time.sleep(2)
browser.find_element_by_css_selector('p[id="next-button"]').click()
time.sleep(5)

result = browser.find_element_by_xpath('//*[@id="resultArea"]/div[3]/h2/strong').text
print(result)
