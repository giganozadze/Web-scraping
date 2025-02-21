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

df = pd.DataFrame(np.nan, index=range(0, 100000), columns=['განცხადება','მისამართი','თარიღი','ნახვა','ID','ფასი $','ფართობი','ფასი $ კვ.მ','ოთახი','საძინებელი','სართული','აღწერა','სივრცე','კეთილმოწყობა'])

website = 'https://www.myhome.ge/ka/s/iyideba-axali-ashenebuli-bina-bakuriani?Keyword=%E1%83%91%E1%83%90%E1%83%99%E1%83%A3%E1%83%A0%E1%83%98%E1%83%90%E1%83%9C%E1%83%98&AdTypeID=1&PrTypeID=1&mapC=41.7510862%2C43.5280065&districts=311913158&cities=311913158&GID=311913158&EstateTypeID=1.2'
path = 'C:/Users/user/Desktop/Projects/Myhome/chromedriver.exe'

driver = webdriver.Chrome(path)
driver.get(website)

nextpage = driver.find_element_by_class_name('next-page')
delay = 12

for u in range(1, 665):

    t = u-int((u-1)/22)*22

    try:
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'card-title')))
    except Exception as e:
        print(e)

    element = driver.find_elements_by_class_name('card-title')[t-1]
    driver.execute_script("arguments[0].click();", element)

    original_window = driver.current_window_handle
    wait = WebDriverWait(driver, delay)
    wait.until(EC.number_of_windows_to_be(2))
    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)
            break

    try:
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#main_block > div.detail-page > div.statement-header > div.info.d-flex.flex-wrap')))
    except Exception as e:
        print(e)

    try:
        announcement = driver.find_element_by_class_name('statement-title').text
        a = announcement.split('\n')
        announcement_title = a[0]
        address = a[1]
        df.iloc[u, 0] = announcement_title
        df.iloc[u, 1] = address
    except:
        announcement = ''
        df.iloc[u, 0] = ''
        df.iloc[u, 1] = ''

    try:
        info = driver.find_element_by_css_selector('#main_block > div.detail-page > div.statement-header > div.info.d-flex.flex-wrap').text
        b = info.split('\n')
        date = b[1]
        views = b[2]
        ann_id = b[3]
        df.iloc[u, 2] = date
        df.iloc[u, 3] = views
        df.iloc[u, 4] = ann_id
    except:
        info = ''
        df.iloc[u, 2] = ''
        df.iloc[u, 3] = ''
        df.iloc[u, 4] = ''

    try:
        info2 = driver.find_element_by_class_name('price-toggler-wrapper').text
        c = info2.split('\n')
        price = c[0]
        area = c[1]
        price_sqm = c[2]

        df.iloc[u, 5] = price
        df.iloc[u, 6] = price_sqm
        df.iloc[u, 7] = area
    except:
        info2 = ''
        df.iloc[u, 5] = ''
        df.iloc[u, 6] = ''
        df.iloc[u, 7] = ''


    try:
        myElem2 = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main_block"]/div[5]/div[4]/div[3]')))
    except Exception as e:
        print(e)

    try:
        size = driver.find_element_by_xpath('//*[@id="main_block"]/div[5]/div[4]/div[1]').text
        d = size.split('\n')
        rooms = d[1]
        df.iloc[u, 8] = rooms
    except:
        size = ''
        df.iloc[u, 8] = ''

    try:
        bedroom = driver.find_element_by_xpath('//*[@id="main_block"]/div[5]/div[4]/div[2]').text
        o = bedroom.split('\n')
        bedrooms = o[0]
        df.iloc[u, 9] = bedrooms
    except:
        bedroom = ''
        df.iloc[u, 9] = bedrooms

    try:
        floor = driver.find_element_by_xpath('//*[@id="main_block"]/div[5]/div[4]/div[3]').text
        f = floor.split('\n')
        floors = f[0]
        df.iloc[u, 10] = floors
    except:
        floor = ''
        df.iloc[u, 10] = ''

    try:
        myElem3 = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#main_block > div.detail-page > div.description > div:nth-child(1) > div > p')))
    except Exception as e:
        print(e)

    try:
        metis_naxva = driver.find_elements_by_class_name('hover-underline')
        for ie in metis_naxva:
            driver.execute_script("arguments[0].click();", ie)
    except:
        print('No such button: metis_naxva')

    try:
        description = driver.find_element_by_css_selector('#main_block > div.detail-page > div.description > div:nth-child(1) > div > p').text
        df.iloc[u, 11] = description
    except:
        description = ''
        df.iloc[u, 11] = ''


    try:
        myElem4 = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#main_block > div.detail-page > div.amenities > div.row > div:nth-child(1) > ul')))
    except Exception as e:
        print(e)

    try:
        sivrce = driver.find_element_by_xpath('//*[@id="main_block"]/div[5]/div[6]/div[1]/div[1]/ul').text
        g = sivrce.split('\n')
        for item in range(1, len(g)):
            j = driver.find_element_by_xpath(
                f'//*[@id="main_block"]/div[5]/div[6]/div[1]/div[1]/ul/li[{item}]/div/span').get_attribute('class')
            if j == 'd-block no':
                g[item] = ''
        g2 = ', '.join(g)
        df.iloc[u, 12] = g2
    except:
        sivrce = ''
        df.iloc[u, 12] = ''

    try:
        ketilmowkoba = driver.find_element_by_xpath('//*[@id="main_block"]/div[5]/div[6]/div[1]/div[2]/ul').text
        h = ketilmowkoba.split('\n')
        for item2 in range(1, len(h)):
            k = driver.find_element_by_xpath(
                f'//*[@id="main_block"]/div[5]/div[6]/div[1]/div[2]/ul/li[{item2}]/div/span').get_attribute('class')
            if k == 'd-block no':
                h[item2] = ''
        h2 = ', '.join(h)
        df.iloc[u, 13] = h2
    except:
        ketilmowkoba = ''
        df.iloc[u, 13] = ''

    driver.close()
    driver.switch_to.window(original_window)

    if u % 22 == 0:
        nextpage.click()
        time.sleep(5)

    try:
        myElem5 = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'next-page')))
    except Exception as e:
        print(e)

    # try:
    #     if u % 3000 == 0:
    #         workbook = xlsxwriter.Workbook('%s.xlsx'%u)
    #         worksheet = workbook.add_worksheet()
    #         ex = df.dropna()
    #         ex.to_excel('%s.xlsx'%u)
    #         workbook.close()
    # except Exception as e:
    #     print(e)


df.drop_duplicates(subset=['განცხადება','მისამართი','თარიღი','ნახვა','ID','ფასი $','ფართობი','ფასი $ კვ.მ'], keep='first', inplace=True, ignore_index=False)



df.to_excel('Bakuriani_Bachana.xlsx')

