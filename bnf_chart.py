# -*- codisng: utf-8 -*-
"""
Created on Sun Feb 16 11:33:05 2020

@author: Nitish Dash
"""


#OPTION CHAIN NSE

from datetime import datetime
from selenium import webdriver
from bs4 import BeautifulSoup as BSoup 
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.firefox.options import Options as FirefoxOptions
import pandas as pd
import flask
import xlsxwriter  
import numpy

def transpose(l1, l2): 
  
    # iterate over list l1 to the length of an item  
    for i in range(len(l1[0])): 
        # print(i) 
        row =[] 
        for item in l1: 
            # appending to new list with values and index positions 
            # i contains index position and item contains values 
            row.append(item[i]) 
        l2.append(row) 
    return l2 

#CONSTANTS
BASE_URL = "https://www1.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?segmentLink=17&instrument=OPTIDX&symbol=BANKNIFTY&date={}"
WEB_DRIVER_PATH = 'C:\devtools\geckodriver.exe'
BINARY = r'C:\Program Files\Mozilla Firefox\firefox.exe'
EXPIRY = "2JUL2020"
FINAL_URL = BASE_URL.format(EXPIRY)
TIMEOUT = 15
now = datetime.now()

cap = DesiredCapabilities().FIREFOX
cap["marionette"] = True

options = FirefoxOptions()
options.binary = BINARY
options.add_argument("--headless")
options.add_argument("--window-size=1920x1080")
driver = webdriver.Firefox(capabilities = cap, executable_path = WEB_DRIVER_PATH, options = options)
driver.get(FINAL_URL)
print('Initiated headless')
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
data_new = []

cols = col_names.find_all('th')
for x in range(1, (len(cols)-1)): 
    column_names.append(cols[x].text)

#Stripper

    
for y in range(0, len(rows)-1):
    row = rows[y]
    cells = row.find_all('td')
    c_oi = int((cells[1].text).replace(',', '').strip()) if (cells[1].text).strip() != '-' else (cells[1].text).strip()
    c_change_oi = int((cells[2].text).replace(',', '').strip()) if (cells[2].text).strip() != '-' else (cells[2].text).strip()
    c_volume = int((cells[3].text).replace(',', '').strip()) if (cells[3].text).strip() != '-' else (cells[3].text).strip()
    c_iv = cells[4].text 
    c_ltp = float((cells[5].text).replace(',', '').strip()) if (cells[5].text).strip() != '-' else (cells[5].text).strip()
    c_change = float((cells[6].text).replace(',', '').strip()) if (cells[6].text).strip() != "-" else (cells[6].text).strip()
    c_bid_quantity = cells[7].text
    c_bid_price = cells[8].text
    c_ask_price = cells[9].text
    c_ask_quantity = cells[10].text 
    strike_price = float((cells[11].text).strip())
    p_bid_quantity = cells[12].text
    p_bid_price = cells[13].text
    p_ask_price = cells[14].text
    p_ask_quantity = cells[15].text 
    p_change = float((cells[16].text).replace(',', '').strip()) if (cells[16].text).strip() != '-' else (cells[16].text).strip()
    p_ltp = float((cells[17].text).replace(',', '').strip()) if (cells[17].text).strip() != '-' else (cells[17].text).strip()
    p_iv = cells[18].text
    p_volume = int((cells[19].text).replace(',', '').strip()) if (cells[19].text).strip() != '-' else (cells[19].text).strip() 
    p_change_oi = int((cells[20].text).replace(',', '').strip()) if (cells[20].text).strip() != '-' else (cells[20].text).strip()
    p_oi = int((cells[21].text).replace(',', '').strip()) if (cells[21].text).strip() != '-' else (cells[21].text).strip() 

    data.append([c_oi,    c_change_oi,    c_volume,    c_iv,    c_ltp,    c_change,    c_bid_quantity,    c_bid_price,    c_ask_price,    c_ask_quantity,    strike_price,    p_bid_quantity,    p_bid_price,    p_ask_price,    p_ask_quantity,    p_change,    p_ltp,    p_iv,    p_volume,    p_change_oi,    p_oi])
    small_data.append([c_oi, c_change_oi, c_volume, c_ltp, c_change, strike_price, p_change, p_ltp, p_volume, p_change_oi, p_oi]);

driver.quit()

#Invert cols and rows

data_new = transpose(small_data, data_new)

#data_new = numpy.transpose(small_data)

dt_string2 = now.strftime("%d.%m.%Y_%H.%M")
file_name2 = dt_string2+"_BNF.xlsx"
workbook = xlsxwriter.Workbook(file_name2)  
worksheet = workbook.add_worksheet()

bold = workbook.add_format({'bold': 1}) 
headings = ['CALL OI', 'COI_CHG','STRIKE', 'POI_CHG', 'PUT OI']
heads = ['CLOSING', 'DATE', 'EXPIRY']
worksheet.write_row('A1', headings, bold)  
worksheet.write('G1', heads[0], bold)
worksheet.write('I1', heads[1], bold)
worksheet.write('K1', heads[2], bold)  
worksheet.write('G2', float(bnf_price), bold)
worksheet.write('I2', time)
worksheet.write('K2', EXPIRY, bold)

for y in range(0, len(small_data)-1):
    worksheet.write_column('A2', data_new[0]) 
    worksheet.write_column('B2', data_new[1])  
    worksheet.write_column('C2', data_new[5])
    worksheet.write_column('D2', data_new[9])  
    worksheet.write_column('E2', data_new[10])
    
chart1 = workbook.add_chart({'type': 'column'}) 
chart2 = workbook.add_chart({'type': 'column'}) 

chart1.add_series({  
    'name':       '=Sheet1!$A$1',  
    'categories': '=Sheet1!$C$59:$C$89',  
    'values':     '=Sheet1!$A$59:$A$89',  
})  
    
chart1.add_series({  
    'name':       '=Sheet1!$E$1',  
    'categories': '=Sheet1!$C$59:$C$89',  
    'values':     '=Sheet1!$E$59:$E$89',  
})     

chart2.add_series({  
    'name':       '=Sheet1!$B$1',  
    'categories': '=Sheet1!$C$59:$C$89',  
    'values':     '=Sheet1!$B$59:$B$89',  
})  
    
chart2.add_series({  
    'name':       '=Sheet1!$D$1',  
    'categories': '=Sheet1!$C$59:$C$89',  
    'values':     '=Sheet1!$D$59:$D$89',  
})         
    
    
chart1.set_title ({'name': 'Open Interest'})  
chart1.set_x_axis({'name': 'Strikes'})     
chart1.set_y_axis({'name': 'Quantity'})

chart2.set_title ({'name': 'Change in Open Interest'})  
chart2.set_x_axis({'name': 'Strikes'})     
chart2.set_y_axis({'name': 'Quantity'})

chart1.set_style(68) 
chart2.set_style(68) 

worksheet.insert_chart('G4', chart1, {'x_scale': 2.8, 'y_scale': 1.3})  
worksheet.insert_chart('G24', chart2, {'x_scale': 2.8, 'y_scale': 2})  


workbook.close()



