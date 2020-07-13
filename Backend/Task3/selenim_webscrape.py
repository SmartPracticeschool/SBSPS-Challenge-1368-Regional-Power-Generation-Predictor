from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter
import requests, sys, re

url = 'https://www.wunderground.com/hourly/in/pune/date/2020-07-08'

path = "C:\webdrivers\chromedriver"
br = webdriver.Chrome(path)
br.get(url)
sopa = BeautifulSoup(br.page_source, 'lxml')
br.quit()

table = sopa.find_all('table')
for i in range(len(table)):
    workbk_out = xlsxwriter.Workbook("test.xlsx")
    sheet_out = workbk_out.add_worksheet()
    table = table[i]
    table_head = table.findAll('th')
    table_data = table.findAll('td')
    output_head = []
    for head in table_head:
        sheet_out.write(head.text.strip())
    enCodeee = '"' + '";"'.join(output_head) + '"'
    enCodeee = re.sub('\s', '', enCodeee) + '\n'
    #out_file.write(enCodeee.encode(encoding='UTF-8'))
    output_rows = []
    files = table.findAll('tr')
    for j in range(1, len(files)):
        table_row = files[j]
        columns = table_row.findAll('td')
        output_row = []
        for column in columns:
            output_row.write(column.text.strip())
        file = '"' + '";"'.join(output_row) + '"'
        file = re.sub('\s', '', file) + '\n'
        out_file.write(file.encode(encoding='UTF-8'))

    workbk_out.close()
