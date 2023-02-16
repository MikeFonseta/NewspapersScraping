from time import sleep
import signal
import os
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from openpyxl import Workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#Recensioni totali estratte
tot_reviews = 0
page = 0
max_reviews = 20
#Apertura foglio excel
wb = Workbook()
ws = wb.active

#Opzioni per il webdriver
options = EdgeOptions()
options.add_argument("start-maximized")
options.add_experimental_option('excludeSwitches', ['enable-logging'])


#Creazione webdriver           In questo caso Edge
driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()),options=options)


def analisi(query="",url=""):
    global tot_reviews,wb,ws,driver,page

    
    #Caricamento pagina
    driver.get(url)
    print(driver.find_element(By.CLASS_NAME, 'biGQs._P.fiohW.eIegw').text)
    #Ricerca

    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'onetrust-accept-btn-handler'))).click()

    searchFiled = driver.find_element(By.CLASS_NAME,'UfnDM.z.w.aThUm.IiSKr')
    searchFiled.send_keys(query)
    sleep(5)
    page = 1
    while tot_reviews <= max_reviews:

        reviews = None
        title = None
        decription = None
        date = None
        link = None

        reviews = driver.find_element(By.CLASS_NAME, 'LbPSX')

        if(reviews is not None):
            
            title = WebDriverWait(reviews, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,'biGQs._P.fiohW.qWPrE.ncFvv.fOtGX')))
            decription = WebDriverWait(reviews, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,'_T.FKffI')))
            date = WebDriverWait(reviews, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,'RpeCd')))
            link = WebDriverWait(reviews, 10).until(EC.visibility_of_all_elements_located((By.CLASS_NAME,'BMQDV._F.G-.wSSLS.SwZTJ.FGwzt.ukgoS')))

            offSet = 0

            # for l in link:
            #     print(l.get_attribute('href'))

            for i in range(0,len(title)):
                # print(title[i].text)
                # print(decription[i].text)
                # print(date[i].text)
                # print(link[offSet+1].get_attribute('href'))
                # print("\n")
                ws.append((title[i].text, decription[i].text , date[i].text, link[offSet+1].get_attribute('href')))
                tot_reviews += 1
                wb.save('Tripadvisor.xlsx')
                offSet+=2

            print("Pagina:",page," | ","Recensioni trovate: " + str(tot_reviews)+" | " + "Recensioni salvate: " + str(tot_reviews))

            if(tot_reviews <= max_reviews):
                elem = driver.find_element(By.CLASS_NAME, 'UCacc')
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
    print("Scraping interrotto \n" + str(tot_reviews) + " recensioni salvate \nAttendere terminazione script.")
    wb.save(name+'.xlsx')
    driver.quit()


try:
    #Signal per identificare quando lo script viene fermato con CTRL+C
    signal.signal(signal.SIGINT, signal.default_int_handler)
    analisi(query="Napoli ", url="https://www.tripadvisor.it/Attraction_Review-g187785-d3171016-Reviews-Napoli_Sotterranea-Naples_Province_of_Naples_Campania.html")
    driver.quit()
except KeyboardInterrupt:
    os._exit(0)

