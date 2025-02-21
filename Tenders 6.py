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
from selenium.webdriver.common.by import By
import xlsxwriter

os.chdir('C:/Users/gigan/OneDrive/Desktop/Tenders')
df = pd.DataFrame(np.nan, index=range(0, 5000), columns=['შესყიდვის ტიპი','განცხადების ნომერი','შესყიდვის სტატუსი','შემსყიდველი','შესყიდვის გამოცხადების თარიღი','წინადადებების მიღება იწყება','წინადადებების მიღება მთავრდება','შესყიდვის სავარაუდო ღირებულება','წინადადება წარმოდგენილი უნდა იყოს','შესყიდვის კატეგორია','კლასიფიკატორის კოდები','მოწოდების ვადა','დამატებითი ინფორმაცია','შესყიდვის რაოდენობა ან მოცულობა','შეთავაზების ფასის კლების ბიჯი','გარანტიის ოდენობა','გარანტიის მოქმედების ვადა','ქრონოლოგია','შეთავაზებები','ხელშეკრულება'])

#####Indicate criteria and search#####
driver = webdriver.Chrome()
driver.get("https://tenders.procurement.gov.ge/public/?lang=ge")

element = driver.find_element(By.XPATH, 'html/body/div[2]/div[3]/div[2]/div/span/button[4]/span[1]')
driver.execute_script("arguments[0].click();", element)

for i in range(1,159):
    t = i-int((i-1)/4)*4
    start_time = time.time()

    ########Open main page#########

    delay = 5
    try:
        s = time.time()
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[@id="container"]/div[@id="content"]/div[2]/table/tbody/tr[4]')))
        print("Page is ready! მთავარი გვერდი")
        print("--- %s მთავარი გვერდის დრო ---" % (time.time() - s))
    except Exception as e:
        print("%s მთავარი გვერდი" % e)

    element = driver.find_element(By.XPATH, '/html/body/div[@id="container"]/div[@id="content"]/div[2]/table/tbody/tr[%s]'% t)
    driver.execute_script("arguments[0].click();", element)

    delay = 5
    try:
        s=time.time()
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[1]/td[2]')))
        print("Page is ready! შიდა პირველი გვერდი")
        print("--- %s შიდა პირველი გვერდი---" % (time.time() - s))
    except Exception as e:
        print("%s შიდა პირველი გვერდი" % e)

    ###შესყიდვის ტიპი###
    try:
        she=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(1) > td:nth-child(2)').text
        #she=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[1]/td[2]').text
        df.iloc[i,0]=she
    except Exception as e:
        df.iloc[i,0] = e

    ##განცხადების ნომერი#
    try:
        idg=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(2) > td:nth-child(2)').text
        #idg=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[2]/td[2]').text
        df.iloc[i,1]=idg
    except Exception as e:
        df.iloc[i,1]=e

    ##შესყიდვის სტატუსი#
    try:
        st=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(3) > td:nth-child(2)').text
        #st=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[3]/td[2]').text
        df.iloc[i,2]=st
    except Exception as e:
        df.iloc[i,2]=e

    ####შემსყიდველი#####
    try:
        buy=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(4) > td:nth-child(2)').text
        #buy=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[4]/td[2]').text
        df.iloc[i,3]=buy
    except Exception as e:
        df.iloc[i,3]=e

    ####შესყიდვის გამოცხადების თარიღი#####
    try:
        dat=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(5) > td:nth-child(2)').text
        #dat=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[5]/td[2]').text
        df.iloc[i,4]=dat
    except Exception as e:
        df.iloc[i,4]=e

    ####წინადადებების მიღება იწყება#####
    try:
        star=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(6) > td:nth-child(2)').text
        #star=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[6]/td[2]').text
        df.iloc[i,5]=star
    except Exception as e:
        df.iloc[i,5]=e

    ###წინადადებების მიღება მთავრდება###
    try:
        ded=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(7) > td:nth-child(2)').text
        #ded=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[7]/td[2]').text
        df.iloc[i,6]=ded
    except Exception as e:
        df.iloc[i,6]=e

    ##შესყიდვის სავარაუდო ღირებულება##
    try:
        pri=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(8) > td:nth-child(2)').text
        #pri=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[8]/td[2]').text
        df.iloc[i,7]=pri
    except Exception as e:
        df.iloc[i,7]=e

    ##წინადადება წარმოდგენილი უნდა იყოს##
    try:
        vat=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(9) > td:nth-child(2)').text
        #vat=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[9]/td[2]').text
        df.iloc[i,8]=vat
    except Exception as e:
        df.iloc[i,8]=e

    ########შესყიდვის კატეგორია############
    try:
        cate=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(10) > td.subject_name').text
        #cate=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[10]/td[2]').text
        df.iloc[i,9]=cate
    except Exception as e:
        df.iloc[i,9]=e

    ########კლასიფიკატორის კოდები########
    try:
        clas=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(12)').text
        #clas=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[12]').text
        df.iloc[i,10]=clas
    except Exception as e:
        df.iloc[i,10]=e

    ############მოწოდების ვადა############
    try:
        tim=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(13) > td:nth-child(2)').text
        #tim=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[13]/td[2]').text
        df.iloc[i,11]=tim
    except Exception as e:
        df.iloc[i,11]=e

    #########დამატებითი ინფორმაცია########
    try:
        ainfo=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(15)').text
        #ainfo=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[15]').text
        df.iloc[i,12]=ainfo
    except Exception as e:
        df.iloc[i,12]=e

    ###შესყიდვის რაოდენობა ან მოცულობა####
    try:
        q=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(16) > td:nth-child(2)').text
        #q=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[16]/td[2]').text
        df.iloc[i,13]=q
    except Exception as e:
        df.iloc[i,13]=e

    ####შეთავაზების ფასის კლების ბიჯი######
    try:
        step=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(17) > td:nth-child(2)').text
        #step=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[17]/td[2]').text
        df.iloc[i,14]=step
    except Exception as e:
        df.iloc[i,14]=e

    ###########გარანტიის ოდენობა##########
    try:
        gur=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(18) > td:nth-child(2)').text
        #gur=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[18]/td[2]').text
        df.iloc[i,15]=gur
    except Exception as e:
        df.iloc[i,15]=e

    #####გარანტიის მოქმედების ვადა#########
    try:
        tgur=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(19) > td:nth-child(2)').text
        #tgur=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[19]/td[2]').text
        df.iloc[i,16]=tgur
    except Exception as e:
        df.iloc[i,16]=e

    # #####add_1#########
    # try:
    #     tgur=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(11) > td:nth-child(2)').text
    #     #tgur=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[19]/td[2]').text
    #     df.iloc[i,20]=tgur
    # except Exception as e:
    #     df.iloc[i,20]=e
    #
    # #####add_2#########
    # try:
    #     tgur=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(12) > td:nth-child(2)').text
    #     #tgur=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[19]/td[2]').text
    #     df.iloc[i,21]=tgur
    # except Exception as e:
    #     df.iloc[i,21]=e
    #
    # #####add_3#########
    # try:
    #     tgur=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(13) > td:nth-child(2)').text
    #     #tgur=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[19]/td[2]').text
    #     df.iloc[i,22]=tgur
    # except Exception as e:
    #     df.iloc[i,22]=e
    #
    # #####add_4#########
    # try:
    #     tgur=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(14) > td:nth-child(2)').text
    #     #tgur=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[19]/td[2]').text
    #     df.iloc[i,23]=tgur
    # except Exception as e:
    #     df.iloc[i,23]=e
    #
    # #####add_5#########
    # try:
    #     cate=driver.find_element(By.CSS_SELECTOR, '##print_area > table > tbody > tr:nth-child(15)').text
    #     #cate=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[10]/td[2]').text
    #     df.iloc[i,24]=cate
    # except Exception as e:
    #     df.iloc[i,24]=e
    #
    # #####add_6#########
    # try:
    #     cate=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(13)').text
    #     #cate=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[10]/td[2]').text
    #     df.iloc[i,25]=cate
    # except Exception as e:
    #     df.iloc[i,25]=e
    #
    # #####add_7#########
    # try:
    #     cate=driver.find_element(By.CSS_SELECTOR, '#print_area > table > tbody > tr:nth-child(20) > td:nth-child(2)').text
    #     #cate=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[1]/table/tbody/tr[10]/td[2]').text
    #     df.iloc[i,26]=cate
    # except Exception as e:
    #     df.iloc[i,26]=e

    delay = 5
    try:
        s = time.time()
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[@class="pad4px"]/span/button')))
        print("Page is ready! ქრონოლოგია")
        print("--- %s შიდა პირველი გვერდი---" % (time.time() - s))
    except Exception as e:
        print("%s  შიდა პირველი გვერდი" % e)

    element = driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[@class="pad4px"]/span/button')
    driver.execute_script("arguments[0].click();", element)
    # qro = driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[@class="pad4px"]/span/button')
    # qro.click()
    delay = 5
    try:
        s = time.time()
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div/div/div[@id="history"]/div/table')))
        print("Page is ready! ქრონოლოგია გახსნილია")
        print("--- %s ქრონოლოგია გახსნილია---" % (time.time() - s))
    except Exception as e:
        print("%s  ქრონოლოგია გახსნილია" % e)
    try:
        qron=driver.find_element(By.XPATH, '//p[contains(string(), "ქრონოლოგია")]/following::*[1]').text
        df.iloc[i,17]=qron
    except Exception as e:
        df.iloc[i,17]=e

    #####შეთავაზებები#####
    element = driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[@id="application_tabs"]/ul/li[3]/a')
    driver.execute_script("arguments[0].click();", element)
    # off=driver.find_element(By.XPATH, 'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[@id="application_tabs"]/ul/li[3]')
    # off.click()

    delay = 5
    try:
        s = time.time()
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH,'html/body//div[@id="container"]/div[@id="content"]//div[3]/div[2]/div[3]/div/table/tbody')))
        print("Page is ready! შეთავაზების გვერდზე გადასვლა")
        print("--- %s შეთავაზების გვერდზე გადასვლა---" % (time.time() - s))
    except Exception as e:
        print("%s  შეთავაზების გვერდზე გადასვლა" % e)
    try:
        part=driver.find_element(By.XPATH, '//*[@id="app_bids"]/table').text
        if part=='':
            df.iloc[i, 18] = "None"
        else:
            df.iloc[i, 18] = part
    except Exception as e:
        df.iloc[i, 18] = e

    #####ხელშეკრულება#####
    try:
        element = driver.find_element(By.XPATH, '/html/body/div[2]/div[3]/div[3]/div[2]/ul/li[5]/a')
        driver.execute_script("arguments[0].click();", element)
    except Exception as e:
        print("ხელშეკრულება არ დადებულა")

    delay = 5
    try:
        s = time.time()
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="agency_docs"]/div[1]')))
        print("Page is ready! ხელშეკრულების გვერდზე გადასვლა")
        print("--- %s ხელშეკრულების გვერდზე გადასვლა---" % (time.time() - s))
    except Exception as e:
        print("%s  ხელშეკრულების გვერდზე გადასვლა" % e)
    try:
        part = driver.find_element(By.XPATH, '//*[@id="agency_docs"]/div[1]/table').text
        if part == '':
            df.iloc[i, 19] = "None"
        else:
            df.iloc[i, 19] = part
    except Exception as e:
        df.iloc[i, 19] = e


    element = driver.find_element(By.XPATH, 'html/body/div[@id="container"]/div[@id="content"]/div[3]/div[1]/button/span[2]')
    driver.execute_script("arguments[0].click();", element)

    if i % 4 == 0:
        element = driver.find_element(By.XPATH, 'html/body/div[2]/div[3]/div[2]/div/span/button[4]/span[1]')
        driver.execute_script("arguments[0].click();", element)
        time.sleep(3)

    if i % 3000 == 0:
        workbook = xlsxwriter.Workbook('%s.xlsx'% i )
        worksheet = workbook.add_worksheet()
        ex = df.dropna(subset=["შესყიდვის ტიპი"])
        ex.to_excel('%s.xlsx'%i)
        workbook.close()
    print("--- %s seconds ---" % (time.time() - start_time))

df.dropna(how='all', inplace=True)
df.to_excel('data_tenders_142.xlsx')
