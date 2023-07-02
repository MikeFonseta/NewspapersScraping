from time import sleep
import signal
import os
from openpyxl import Workbook
from datetime import datetime
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

months = {'Jan.': '01',
        'Feb.': '02',
        'March': '03',
        'April': '04',
        'May': '05',
        'June': '06',
        'July': '07',
        'Aug.': '08',
        'Sept.': '09',
        'Oct.': '10',
        'Nov.': '11',
        'Dec.': '12',}
articles_found = 0
tot_articles = 0
wb = Workbook()
ws = wb.active
url = "https://www.nytimes.com/search?dropmab=false&query=uffizi&sort=oldest"
options = Options()
options.add_argument("start-maximized")
options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options) 


def analisi():
    global wb,ws,tot_articles,articles_found,url
    
    #Numero massimo di articoli da prendere   
    max_articles = 100
    

    driver.get(url)

    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'css-ni9it0.e1j3jvdr0'))).click()
    sleep(15)
    driver.find_element(By.NAME, 'email').send_keys('beccherlepaola@yahoo.it')
    driver.find_element(By.CLASS_NAME, 'css-1i3jzoq-buttonBox-buttonBox-primaryButton-primaryButton-Button').click()
    sleep(2)
    driver.find_element(By.NAME, 'password').send_keys('Scraping2023')
    driver.find_element(By.CLASS_NAME, 'css-1i3jzoq-buttonBox-buttonBox-primaryButton-primaryButton-Button').click()
    sleep(15)


    while(tot_articles < max_articles):
        
        while len(driver.find_elements(By.CLASS_NAME, 'css-1l4w6pd')) < max_articles:
            sleep(2)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
            driver.find_element(By.CLASS_NAME, 'css-vsuiox').find_element(By.CSS_SELECTOR, 'button').click()
            sleep(2)

        articles_found = driver.find_elements(By.CLASS_NAME, 'css-1l4w6pd')

        index = tot_articles
        print("INDEX BEFORE FOR =>", index,len(articles_found))
        for index in range(tot_articles,len(articles_found)):
            print("===",index)
            driver.execute_script("arguments[0].scrollIntoView(true);", articles_found[index])

            date = articles_found[index].find_element(By.CLASS_NAME,'css-17ubb9w').text

            date = date.split(' ')
            date[1] = date[1].replace(',','')
            date = ("0"+str(date[1]) if (int(date[1]) < 10) else str(date[1])) + "/" +  months[date[0]] + "/" + date[2]
            title = articles_found[index].find_element(By.CLASS_NAME,'css-2fgx4k').text
            link = articles_found[index].find_element(By.CSS_SELECTOR,'a').get_attribute('href')

            main_window = driver.current_window_handle
            text = ""
            if link:
                driver.execute_script("window.open('"+link+"', 'new_window')")
                driver.switch_to.window(driver.window_handles[1])
                sleep(1)
                try:
                    div = driver.find_element(By.CLASS_NAME, 'css-53u6y8')
                    ps = div.find_elements(By.CSS_SELECTOR, 'p')
                    for p in ps:
                        text += p.text
                except:
                    text = driver.find_element(By.CLASS_NAME,'css-6hi8ev').text
                    link = driver.find_element(By.CLASS_NAME,'css-1382yzd').get_attribute('href')


            driver.close()
            sleep(2)
            driver.switch_to.window(main_window)
            ws.append((title, date, link, text))
            wb.save('NewYorkTimes.xlsx')
            tot_articles+=1
        # sleep(2)
        # driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
        # driver.find_element(By.CLASS_NAME, 'css-vsuiox').find_element(By.CSS_SELECTOR, 'button').click()
        # sleep(2)
        print(tot_articles, " articoli salvati")
        
        # print(len(driver.find_elements(By.CLASS_NAME, 'css-1l4w6pd')))
    save()


def save():
    global tot_articles,wb,driver
    print("Scraping interrotto \n" + str(tot_articles) + " articoli salvati \nAttendere terminazione script.")
    wb.save('NewYorkTimes.xlsx')
    driver.quit()


try:
    #Signal per identificare quando lo script viene fermato con CTRL+C
    signal.signal(signal.SIGINT, signal.default_int_handler)
    analisi()
except KeyboardInterrupt:
    os._exit(0)


