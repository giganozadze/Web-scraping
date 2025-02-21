from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import os
from selenium.webdriver.support.ui import Select
import pandas as pd
import numpy as np
from time import sleep
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import xlsxwriter

os.chdir('C:/Users/User/Desktop/Projects/TendersF')
df = pd.DataFrame(np.nan, index=range(0, 1000), columns=['Name', 'Date'])

driver = webdriver.Chrome()
driver.get("https://galtandtaggart.com/en/media/blog")

for i in range(1, 90):
    t = i-int((i-1)/9)*9

    delay = 3
    try:
        s = time.time()
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="news-outer-wrapper"]/div/app-news-press-realese/div/div/div[2]/div[%s]/app-news-item-teaser/a/div[2]/div/a'% t)))
        print("Page is ready! მთავარი გვერდი")
        print("--- %s მთავარი გვერდის დრო ---" % (time.time() - s))
    except Exception as e:
        print("%s მთავარი გვერდი" % e)

    element = driver.find_element_by_xpath('//*[@id="news-outer-wrapper"]/div/app-news-press-realese/div/div/div[2]/div[%s]/app-news-item-teaser/a/div[2]/div/a'% t)
    driver.execute_script("arguments[0].click();", element)

    delay = 3

    # try:
    #     s=time.time()
    #     myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div[2]/div[3]/div/section/div[1]/div/div[1]/h1')))
    #     print("Page is ready! შიდა პირველი გვერდი")
    #     print("--- %s შიდა პირველი გვერდი---" % (time.time() - s))
    # except Exception as e:
    #     print("%s შიდა პირველი გვერდი" % e)

    ###დასახელება###
    try:
        name = driver.find_element_by_xpath('//*[@id="app"]/div/div[2]/div[3]/div/section/div[1]/div/div[1]/h1').text
        #she=driver.find_element_by_xpath('html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[1]/td[2]').text
        df.iloc[i, 0] = name
    except Exception as e:
        df.iloc[i, 0] = e

    ###თარიღი###
    try:
        date = driver.find_element_by_xpath('//*[@id="app"]/div/div[2]/div[3]/div/section/div[1]/div/div[1]/span').text
        #idg=driver.find_element_by_xpath('html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[2]/td[2]').text
        df.iloc[i, 1] = date
    except Exception as e:
        df.iloc[i, 1] = e

    driver.back()

    delay = 3
    try:
        s = time.time()
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div[2]/div[3]/div[2]/section/div[3]/div/button[2]')))
        print("Page is ready! მთავარი გვერდი")
        print("--- %s მთავარი გვერდის დრო ---" % (time.time() - s))
    except Exception as e:
        print("%s მთავარი გვერდი" % e)

    if i % 9 == 0:
        element = driver.find_element_by_xpath('//*[@id="app"]/div/div[2]/div[3]/div[2]/section/div[3]/div/button[2]')
        driver.execute_script("arguments[0].click();", element)
        time.sleep(2)

df.dropna(how='all', inplace=True)
df.to_excel('tbc_reports.xlsx')
