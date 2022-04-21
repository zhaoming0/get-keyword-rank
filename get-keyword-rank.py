# pip install -i https://pypi.tuna.tsinghua.edu.cn/simple requests pandas bs4 pyquery openpyxl selenium pillow 
import requests
import time
import re
import pandas as pd
from PIL import Image
from bs4 import BeautifulSoup
from pyquery import PyQuery as pq
import openpyxl
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import date
import datetime
import csv
import sys
import string
import os

excelContent={}

def readExcel(file=sys.argv[1]):
    df = pd.DataFrame(pd.read_excel(sys.argv[1],keep_default_na=False))
    asinlist = []
    keylist = []

    for indexs in df.index:
        row = (df.loc[indexs].values)

        ASINs = str(row[0])
        ASINs=ASINs.strip()

        KEYS = str(row[1])
        KEYS=KEYS.strip()

        asinlist.append(ASINs)
        keylist.append(KEYS)
    # remove repeat & null
    asinlist = ([i for i in asinlist if i !=''])
    asinlist = [i for n, i in enumerate(asinlist) if i not in asinlist[:n]]
    keylist = ([i for i in keylist if i !=''])
    keylist = [i for n, i in enumerate(keylist) if i not in keylist[:n]]


    excelContent['key'] = keylist
    excelContent['asin'] = asinlist
    return excelContent

    
# US JP
def select_area_code(driver, code):
    while True:
        try:
            driver.find_element_by_id('glow-ingress-line1').click()
            time.sleep(2)
        except Exception as e:
            driver.refresh()
            time.sleep(10)
            continue
        try:
            driver.find_element_by_id("GLUXChangePostalCodeLink").click()
            time.sleep(1)
        except:
            pass
        try:
            driver.find_element_by_id('GLUXZipUpdateInput').send_keys(code)
            time.sleep(1)
            break
        except Exception as NoSuchElementException:
            try:
                driver.find_element_by_id('GLUXZipUpdateInput_0').send_keys(
                    code.split('-')[0])
                time.sleep(1)
                driver.find_element_by_id('GLUXZipUpdateInput_1').send_keys(
                    code.split('-')[1])
                time.sleep(1)
                break
            except Exception as NoSuchElementException:
                driver.refresh()
                time.sleep(5)
                continue
        print("input area code again")
    driver.find_element_by_id('GLUXZipUpdate').click()
    time.sleep(1)
    driver.refresh()
    time.sleep(1)
    

def main(keyword, page, country='US',flag=0):
    asin_lst = []
    pg = []
    pos = []
    sp = []
    kw_lst = []
    options = webdriver.ChromeOptions()
    options.add_argument('--disable-gpu')
    options.add_argument("disable-web-security")
    options.add_argument('disable-infobars')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    #disable picture
    prefs = {'profile.managed_default_content_settings.images': 2}
    options.add_experimental_option('prefs',prefs)
    #disable picture
    options.add_argument('log-level=3')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--disable-images')

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)
    driver.maximize_window()

    if country == 'JP':
        search_page_url = 'https://www.amazon.co.jp/s?k={}&page={}'
        area = "426-0009"
    elif country == 'US':
        search_page_url = 'https://www.amazon.com/s?k={}&page={}'
        area = "10001"


    for kw in keyword:
        for i in range(1, page + 1):
            kw_mdf = '+'.join(kw.split(' '))
            print("正在爬取", search_page_url.format(kw_mdf, i))
            driver.get(search_page_url.format(kw_mdf, i))
            time.sleep(2)

            if flag == 0:
                select_area_code(driver, area)
                flag = 1

            wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div.s-result-list"
                     )))  #显示wait 20s until Keyword search result list page
            html = pq(driver.page_source)
            items = html('div[data-asin^=B0]').items()
            
            for num, item in enumerate(items):
                kw_lst.append(kw)
                pos.append(num+1)
                pg.append(i)
                asin_lst.append(item.attr("data-asin"))
                sp.append(item(".s-label-popover-default span.a-color-secondary").text())

    res = pd.DataFrame({
        'KeyWord': kw_lst,
        'Position': pos,
        'Page': pg,
        'Asin': asin_lst,
        'SP': sp
    })
    driver.close()
    print("爬取结束")
    return res


def find(df, asin_lst):
    asin_ser = pd.Series(asin_lst, name='Asin')
    res = pd.merge(df, asin_ser, on='Asin')
    return res

nowTime = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-')

#read excel file
excelContent = readExcel()

RUNNING = main(excelContent['key'], 5, sys.argv[2])
RESULT = find(RUNNING, excelContent['asin'])

# save to local
RESULT.to_csv(str(sys.argv[2]) + '-' + nowTime + 'result.csv')
print('result save to :' + str(sys.argv[2]) + '-' + nowTime + 'result.csv')

# python get-keyword-rank.py test.xlsx JP
# test.xlsx --ASIN KEY

