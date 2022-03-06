import requests
import json
import xlwt, xlrd
from bs4 import BeautifulSoup
from selenium import webdriver
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import time
import json


def browser():
    chrome_options = Options()
    driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), chrome_options=chrome_options)
    return driver

def get_search():
    rb = xlrd.open_workbook('поиск.xls',formatting_info=True)
    sheet = rb.sheet_by_index(0)

    data = []
    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        data.append(row[0])
    
    return data

def save_excel(out):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('WB')

    for x,i in enumerate(out):
        ws.write(0+x, 0,i[0])
        ws.write(0+x, 1,i[1])

    wb.save('WB_output.xls')
    return


def main_itter(data, driver):
    out = []
    itter = 0
    for row in data:
        itter += 1
        if itter >= 100:
            print('Restart browser')
            driver.close()
            driver = browser()
            itter = 0

        text = row.replace(' ', '+')
        if text == '':
            continue

        driver.get(f'https://www.wildberries.ru/catalog/0/search.aspx?search={text}')
        start = time.time()
        status = True
        while True:
            if time.time() - start > 35:
                status = False
                break

            if time.time() - start > 25:
                driver.get(f'https://www.wildberries.ru/catalog/0/search.aspx?search={text}')

            if driver.find_elements_by_xpath('//*[@id="catalog"]/div[1]/div[1]/div/span/span[1]'):
                time.sleep(2)
                if driver.find_element_by_xpath('//*[@id="catalog"]/div[1]/div[1]/div/span/span[1]').text.strip().replace(' ', '').isdigit():
                    break

        if status == False:
            continue
        count = driver.find_element_by_xpath('//*[@id="catalog"]/div[1]/div[1]/div/span/span[1]').text
        count = count.strip().replace(' ', '')
        out.append([row, count])

        with open('db.json', 'w') as f:
            json.dump(out, f, sort_keys=False, indent=4, ensure_ascii=False)
    return out, driver


def start():
    driver = browser()

    data = get_search()
    out, driver = main_itter(data, driver)

    driver.close()
    save_excel(out)
start()












