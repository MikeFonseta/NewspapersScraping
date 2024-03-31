from time import sleep
import sys
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import signal
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from datetime import datetime,timezone
import time
from pandas.tseries.offsets import DateOffset

EMAIL = "sivaxer157@bizatop.com"
PASSWORD = "PasswordApi"
USERNAME = "ProvaA6860"
data = []
currentIndex = 0
linkFound = []

def analisi(allOfTheseWords,thisExactPhrase="",anyOfTheseWords="",noneOfTheseWords="",theseHashtags="",language="en",fromDate="31/12/2018",toDate="31/12/2023",replies=False,max=-1,hide=False,latest=False):
    global currentIndex,data


    if toDate == "":
        toDate = datetime.today().strftime('%d-%m-%Y')

    dates = pd.date_range(start=fromDate, end=toDate, freq=DateOffset(months=1)).astype(str)
    dates = ['#'.join(x) for x in zip(dates[: -1], dates[1: ])]


    options = Options()
    if hide == True:
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)

    if hide == False: 
        driver.maximize_window()

    driver.get("https://twitter.com/search-advanced")

    WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.CLASS_NAME, 'r-30o5oe.r-1dz5y72.r-13qz1uu.r-1niwhzg.r-17gur6a.r-1yadl64.r-deolkf.r-homxoj.r-poiln3.r-7cikom.r-1ny4l3l.r-t60dpp.r-fdjqy7'))).send_keys(EMAIL)
    WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.CLASS_NAME, 'css-175oi2r.r-sdzlij.r-1phboty.r-rs99b7.r-lrvibr.r-ywje51.r-usiww2.r-13qz1uu.r-2yi16.r-1qi8awa.r-ymttw5.r-1loqt21.r-o7ynqc.r-6416eg.r-1ny4l3l'))).click()

    passwordField = False

    while not passwordField:
        try:
            WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.NAME, 'password'))).send_keys(PASSWORD)
            WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.CLASS_NAME, 'css-175oi2r.r-sdzlij.r-1phboty.r-rs99b7.r-lrvibr.r-19yznuf.r-64el8z.r-1dye5f7.r-1loqt21.r-o7ynqc.r-6416eg.r-1ny4l3l'))).click()
            passwordField = True
        except:
            WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.NAME, 'text'))).send_keys(USERNAME)
            WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.CLASS_NAME, 'css-175oi2r.r-sdzlij.r-1phboty.r-rs99b7.r-lrvibr.r-19yznuf.r-64el8z.r-1dye5f7.r-1loqt21.r-o7ynqc.r-6416eg.r-1ny4l3l'))).click()
            pass
    

    for x in dates:

        sleep(3)
        driver.get("https://twitter.com/search-advanced")
        driver.refresh()

        twoDates = x.split('#')
        fromDate = twoDates[0].split('-')
        fromDate = fromDate[2] + "/" + fromDate[1] + "/" + fromDate[0]
        toDate = twoDates[1].split('-')
        toDate = toDate[2] + "/" + toDate[1] + "/" + toDate[0]
        print("Prendo tweets da " + fromDate + " a " +toDate)

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, 'allOfTheseWords'))).send_keys(allOfTheseWords)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, 'thisExactPhrase'))).send_keys(thisExactPhrase)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, 'anyOfTheseWords'))).send_keys(anyOfTheseWords)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, 'noneOfTheseWords'))).send_keys(noneOfTheseWords)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, 'theseHashtags'))).send_keys(theseHashtags)

        try:
            Select(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'SELECTOR_1')))).select_by_value(language)
        except:
            pass

        if replies == False:
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'r-30o5oe.r-1p0dtai.r-1pi2tsx.r-1d2f490.r-crgep1.r-t60dpp.r-u8s1d.r-zchlnj.r-ipm5af.r-13qz1uu.r-1ei5mc7'))).click()

        fromDate = fromDate.split('/')

        if len(fromDate) == 3:
            Select(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'SELECTOR_2')))).select_by_value(fromDate[1].strip("0"))
            Select(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'SELECTOR_3')))).select_by_value(fromDate[0].strip("0"))
            Select(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'SELECTOR_4')))).select_by_value(fromDate[2])

        toDateSplitted = toDate.split('/')

        if len(toDateSplitted) == 3:
            Select(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'SELECTOR_5')))).select_by_value(toDateSplitted[1].strip("0"))
            Select(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'SELECTOR_6')))).select_by_value(toDateSplitted[0].strip("0"))
            Select(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'SELECTOR_7')))).select_by_value(toDateSplitted[2])
        
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, 'allOfTheseWords'))).send_keys(Keys.ENTER)

        sleep(3)

        if latest == True:
            driver.get(driver.current_url+"&f=live")

        lastElem = None
        noData = False

        while not noData and (currentIndex < max or max == -1):
            try:
                tweets = WebDriverWait(driver, 3).until(EC.visibility_of_all_elements_located((By.TAG_NAME, 'article')))
                
                

                for x in range(0, len(tweets)):
                    try:
                        utente = WebDriverWait(tweets[x], 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'css-1qaijid.r-bcqeeo.r-qvutc0.r-poiln3'))).text
                        contenuto = WebDriverWait(tweets[x], 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'css-1rynq56.r-8akbws.r-krxsd3.r-dnmrzs.r-1udh08x.r-bcqeeo.r-qvutc0.r-37j5jr.r-a023e6.r-rjixqe.r-16dba41.r-bnwqim'))).text
                        dateTime = WebDriverWait(tweets[x], 10).until(EC.visibility_of_element_located((By.TAG_NAME, 'time'))).get_attribute("datetime")
                        link =  WebDriverWait(tweets[x], 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'css-1rynq56.r-bcqeeo.r-qvutc0.r-37j5jr.r-a023e6.r-rjixqe.r-16dba41.r-xoduu5.r-1q142lx.r-1w6e6rj.r-9aw3ui.r-3s2u2q.r-1loqt21'))).get_attribute("href")

                        if link not in linkFound:
                            linkFound.append(link)
                        
                            dateTime = datetime.fromisoformat(dateTime[:-1]).astimezone(timezone.utc).strftime('%Y-%m-%d')
                            
                            if (time.strptime(dateTime, '%Y-%m-%d') > time.strptime(toDate, '%d/%m/%Y') if latest == True else time.strptime(dateTime, '%d/%m/%Y') < time.strptime(toDate, '%d/%m/%Y')):
                                noData = True
                            else:
                                data.append([utente, contenuto, dateTime,link])
                                currentIndex+=1
                    except Exception as e:
                        pass
                
                if lastElem == tweets[len(tweets)-1]:
                    noData = True
                    print("Tweets finiti per la ricerca effettuata")
                else:
                    lastElem = tweets[len(tweets)-1]
                    driver.execute_script("arguments[0].scrollIntoView();", lastElem)
                    sleep(2)
                    print(str(currentIndex) + " tweets totali salvti")
                pd.DataFrame(data).to_excel('Twitter.xlsx',index=False, header=False)
            except Exception as e:
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'css-175oi2r.r-sdzlij.r-1phboty.r-rs99b7.r-lrvibr.r-2yi16.r-1qi8awa.r-ymttw5.r-1loqt21.r-o7ynqc.r-6416eg.r-1ny4l3l'))).click()
                sleep(5)
                pd.DataFrame(data).to_excel('Twitter.xlsx',index=False, header=False)
                noData = True
                pass
    signal_handler()


def signal_handler():
    pd.DataFrame(data).to_excel('Twitter.xlsx',index=False, header=False)
    print('Salvataggio tweets recuperati ('+str(currentIndex)+')')
    sys.exit(0)
    
try:
    #Signal per identificare quando lo script viene fermato con CTRL+C
    signal.signal(signal.SIGINT, signal_handler)
    analisi(
    latest=True,
    allOfTheseWords="Uffizi"
    )
except KeyboardInterrupt:
    os._exit(0)