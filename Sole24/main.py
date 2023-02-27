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
from selenium.webdriver.common.by import By

#Numero di pagine    Inserire -1 per non mettere nessun limite
max_page = -1
#Pagina corrente
current_page = 1
#Pagina corrente
starting_page = 1
#Articoli totali estratti
tot_articles = 0

#Apertura foglio excel
wb = Workbook()
ws = wb.active

#Opzioni per il webdriver
options = EdgeOptions()
options.add_argument("start-maximized")
options.add_experimental_option('excludeSwitches', ['enable-logging'])

#Creazione webdriver           In questo caso Edge
driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()),options=options)

def analisi():
    global max_page,current_page,tot_articles,wb,ws,driver

    #Caricamento pagina
    driver.get("https://www.ricerca24.ilsole24ore.com")

    #Cookie
    wait = WebDriverWait(driver, 10)
    wait.until(EC.visibility_of_element_located((By.ID, 'onetrust-accept-btn-handler'))).click()

    #Ricerca
    search = driver.find_element(By.CLASS_NAME, value="input.input--lined")
    search.send_keys("Brescia Cultura")
    driver.find_element(By.CLASS_NAME, "search-submit.search-input-submit").click()
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="mainContent"]/div[1]/div/div[1]/div/div[2]/ul[1]/li[2]/a'))).click()


    #Analisi pagine
    while current_page <= max_page or max_page==-1:

        sleep(2)
        #Caricamento pagina
        height = driver.execute_script("return document.body.scrollHeight")
        for scrol in range(100,height,100):
            driver.execute_script(f"window.scrollTo(0,{scrol})")
            sleep(0.25)

        if(current_page >= starting_page):
    
            #Recupero articoli pagina corrente
            article_container = driver.find_element(By.CLASS_NAME, value="list-lined.list-lined--sep.bytime")
            articles = article_container.find_elements(By.CLASS_NAME,value="list-lined-item")


            #Recupero informazioni di ogni singolo articolo della pagina corrente
            for article in articles:
                driver.execute_script("arguments[0].scrollIntoView();", article)
                title = article.find_element(By.CLASS_NAME, value="aprev-title")
                date = article.find_element(By.CLASS_NAME, value="meta-part.time")
                link = title.find_element(By.CSS_SELECTOR, value='a').get_attribute('href')
                text = ""
                sleep(0.75)

                if(link):
                    try:
                        if (link):
                            main_window = driver.current_window_handle
                            driver.execute_script("window.open('"+link+"', 'new_window')")
                            driver.switch_to.window(driver.window_handles[1])
                            sleep(1)
                            ps = driver.find_elements(By.CLASS_NAME, 'atext')
                            for p in ps:
                                text += p.text
                            driver.close()
                            sleep(1)
                            driver.switch_to.window(main_window)
                    except:
                        text = "Error"

                ws.append((title.text, date.get_attribute('datetime'), link, text))
                tot_articles+=1
                wb.save('Scraping.xlsx')

            print("Pagina:",current_page," | ","Articoli trovati: " + str(len(articles))+" | " + "Articoli salvati: " + str(tot_articles))

        #Carico prossima pagina
        next = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'page-link.next')))
        driver.execute_script("arguments[0].scrollIntoView();", next)
        if next is None:
            save()

        next.click()
        current_page+=1

    #Salvataggio file excel
    driver.quit()

#Funzione utilizzata nel caso lo script venga interrotto con CTRL+C.
def save():
    global tot_articles,wb,driver
    print("Scraping interrotto \n" + str(tot_articles) + " articoli salvati \nAttendere terminazione script.")
    wb.save('Scraping.xlsx')
    driver.quit()


try:
    #Signal per identificare quando lo script viene fermato con CTRL+C
    signal.signal(signal.SIGINT, signal.default_int_handler)
    analisi()
except KeyboardInterrupt:
    save()
    os._exit(0)

