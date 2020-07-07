# -*- codisng: utf-8 -*-
"""
Created on Sun Feb 16 11:33:05 2020

@author: Nitish Dash
"""


#OPTION CHAIN NSE
from datetime import datetime
from selenium import webdriver
from bs4 import BeautifulSoup as BSoup 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import flask

#CONSTANTS
BASE_URL = "https://www1.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?segmentLink=17&instrument=OPTIDX&symbol=BANKNIFTY&date={}"
WEB_DRIVER_PATH = 'C:/devtools/chromedriver.exe'
EXPIRY = "18JUN2020"
FINAL_URL = BASE_URL.format(EXPIRY)
TIMEOUT = 15

option = webdriver.ChromeOptions()
#option.add_argument(' --incognito')
driver = webdriver.Chrome(executable_path = WEB_DRIVER_PATH, options = option)
driver.get(FINAL_URL)

print('200')

try:
    WebDriverWait(driver, TIMEOUT).until(EC.visibility_of_element_located((By.XPATH, "//td[@class='ylwbg']")))
    print("Page successfully loaded")
except TimeoutException:
    print("Timed out waiting for page to load")
    driver.quit()


bs_obj = BSoup(driver.page_source, 'html.parser')
rows = bs_obj.find('table', {'id':'octable'}).find('tbody').find_all('tr')
col_names = bs_obj.find('table', {'id':'octable'}).find('thead').find_all('tr')[1]
current_ticker = bs_obj.find_all('table')[0].find('tbody').find('tr').find_all('td')[1].find('div').find_all('span') 

bnf_price = current_ticker[0].find('b').text
time = current_ticker[1].text

#Cleaning up
bnf_price = bnf_price[10:len(bnf_price)]
time = time[6:(len(time)-4)]

print(float(bnf_price))
print(time)

data = []
small_data = []
column_names = []

cols = col_names.find_all('th')
for x in range(1, (len(cols)-1)): 
    column_names.append(cols[x].text)

for y in range(0, len(rows)-1):
    row = rows[y]
    cells = row.find_all('td')
    c_oi = cells[1].text
    c_change_oi = cells[2].text
    c_volume = cells[3].text
    c_iv = cells[4].text
    c_ltp = cells[5].text
    c_change = cells[6].text
    c_bid_quantity = cells[7].text
    c_bid_price = cells[8].text
    c_ask_price = cells[9].text
    c_ask_quantity = cells[10].text 
    strike_price = cells[11].text
    p_bid_quantity = cells[12].text
    p_bid_price = cells[13].text
    p_ask_price = cells[14].text
    p_ask_quantity = cells[15].text 
    p_change = cells[16].text
    p_ltp = cells[17].text
    p_iv = cells[18].text
    p_volume = cells[19].text
    p_change_oi = cells[20].text
    p_oi = cells[21].text

    data.append([c_oi,    c_change_oi,    c_volume,    c_iv,    c_ltp,    c_change,    c_bid_quantity,    c_bid_price,    c_ask_price,    c_ask_quantity,    strike_price,    p_bid_quantity,    p_bid_price,    p_ask_price,    p_ask_quantity,    p_change,    p_ltp,    p_iv,    p_volume,    p_change_oi,    p_oi])
    small_data.append([c_oi, c_change_oi, c_volume, c_ltp, c_change, strike_price, p_change, p_ltp, p_volume, p_change_oi, p_oi]);

driver.quit()
stock_df = pd.DataFrame(small_data, columns = ["OI", "CHANGE OI", "VOL", "LTP", "CHANGE", "STRIKE", "CHANGE", "LTP", "VOL", "CHANGE OI", "OI"])
#print(stock_df)
stock_df.to_excel("BNF_old.xlsx", sheet_name="sheet_1")



