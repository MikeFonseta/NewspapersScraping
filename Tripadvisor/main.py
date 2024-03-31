from time import sleep
import signal
import os
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.edge.service import Service
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from openpyxl import Workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException 

mesi = {'gen': '01',
        'feb': '02',
        'mar': '03',
        'apr': '04',
        'mag': '05',
        'giu': '06',
        'lug': '07',
        'ago': '08',
        'set': '09',
        'ott': '10',
        'nov': '11',
        'dic': '12',}


#Recensioni totali estratte
tot_reviews = 0
max_reviews = 1000
#Apertura foglio excel
wb = Workbook()
ws = wb.active


#Opzioni per il webdriver
options = EdgeOptions()
options.add_argument("start-maximized")
options.add_experimental_option('excludeSwitches', ['enable-logging'])


#Creazione webdriver           In questo caso Edge
driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()),options=options)


def analisi(query="",url="", language="Italiano", mese=None):
    global tot_reviews,wb,ws,driver,page
    
    page = 0

    #Caricamento pagina
    driver.get(url)

    captcha = 0

    while captcha is not None:
        try:
            print("Controllo presenza captcha...")
            captcha = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH,"/html/body/iframe"))).get_attribute("src")
            if("https://geo.captcha-delivery.com/captcha" in captcha):
                print("Captcha rilevato... Apro nuova sessione")
                raise Exception("Captcha rilevato... Apro nuova sessione")
        except TimeoutException:
            captcha = None
        except Exception:
            driver.close()
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
            driver.get(url)

    titlePage = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'biGQs._P.fiohW.eIegw'))).text
    #Ricerca

    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'onetrust-accept-btn-handler'))).click()

    searchFiled = driver.find_element(By.CLASS_NAME,'UfnDM.w.Q')
    searchFiled.send_keys(query)
    
    #sleep(5)
    WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, 'OKHdJ.z.Pc.PQ.Pp.PD.W._S.Gn.Rd._M.PQFNM.wSSLS')))[1].click()
    
    sleep(1)
 
    if mese != None :
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'OKHdJ.z.Pc.PQ.Pp.PD.W._S.Gn.Rd._M.hzzSG.PQFNM.wSSLS'))).click()
        WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, 'OKHdJ.z.Pc.PQ.Pp.PD.W._S.Gn.Rd._M.xARtZ.uPlAb.hzzSG.PQFNM.wSSLS')))[mese-1].click()

    languages = driver.find_element(By.CLASS_NAME, 'IIbRQ._g.z').find_elements(By.CLASS_NAME, 'whtrm._G.z.u.Pi.PW.Pv.PI._S.Wh.Wc.B-.iRKoF')
    
    for languageChoice in languages: 
        try:
            if language in languageChoice.text:
                languageChoice.click()
                contains = " che contengono " + '"'+ query + '"'
                if not query: 
                    print("Cerco recensioni di " + titlePage + " in lingua " + language)
                else:
                    print("Cerco recensioni di " + titlePage + " in lingua " + language + contains)
        except:
            pass
    
    page = 1
    while tot_reviews < max_reviews:

        reviews = None
        title = None
        decription = None
        date = None
        link = None

        reviews = driver.find_element(By.CLASS_NAME, 'LbPSX')

        if(reviews is not None):
            
            title = WebDriverWait(reviews, 5).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,'biGQs._P.fiohW.qWPrE.ncFvv.fOtGX')))
            decription = WebDriverWait(reviews, 5).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,'_T.FKffI')))
            date = WebDriverWait(reviews, 5).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,'RpeCd')))
            link = WebDriverWait(reviews, 5).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,'BMQDV._F.Gv.wSSLS.SwZTJ.FGwzt.ukgoS')))
            starsList = WebDriverWait(reviews, 5).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, 'UctUV.d.H0')))


            offSet = 0

            for i in range(0,len(title)):

                starValue = 0
                stars = starsList[i].find_elements(By.CSS_SELECTOR, 'path')
                for star in stars:
                    if(star.get_attribute('d') == 'M 12 0C5.388 0 0 5.388 0 12s5.388 12 12 12 12-5.38 12-12c0-6.612-5.38-12-12-12z'):
                        starValue += 1
                
                starText = ""
                if (starValue > 1 or starValue == 0):  
                    starText = str(starValue) + " stelle" 
                else: 
                    starText = str(starValue) + " stella"
                try:
                    ws.append((title[i].text, decription[i].text , "01/"+mesi[date[i].text[:3]]+"/"+date[i].text[4:9], link[offSet+1].get_attribute('href'), starText))
                    tot_reviews += 1
                except:
                    ws.append(('Errore','Errore','Errore','Errore','Errore'))
                
                wb.save('Tripadvisor.xlsx')
                offSet+=2

            print("Pagina:",page," | ","Recensioni trovate: " + str(tot_reviews)+" | " + "Recensioni salvate: " + str(tot_reviews))


            if(tot_reviews < max_reviews):
                elem = driver.find_elements(By.CLASS_NAME, 'UCacc')
                elem = elem[len(elem)-1]
                try:
                    elem.click()
                    page+=1
                    sleep(10)
                except:
                    print('Pagine terminate')
            else:
                save()



# Funzione utilizzata nel caso lo script venga interrotto con CTRL+C.
def save(name="Tripadvisor"):
    global tot_reviews,wb,driver
    print(str(tot_reviews) + " recensioni salvate \nAttendere terminazione script.")
    wb.save(name+'.xlsx')
    driver.quit()


try:
    #Signal per identificare quando lo script viene fermato con CTRL+C
    signal.signal(signal.SIGINT, signal.default_int_handler)
    analisi(query="Florence ", language="Inglese",url="https://www.tripadvisor.it/Attraction_Review-g187895-d191153-Reviews-Gallerie_Degli_Uffizi-Florence_Tuscany.html", mese=2)
    driver.quit()
except KeyboardInterrupt:
    os._exit(0)

