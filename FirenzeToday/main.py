from time import sleep
import signal
import os
from datetime import datetime
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from openpyxl import Workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

filters = {'recenti': '',
        'vecchi': '/sort/oldest',
        'rilevanza': '/sort/best',
        'letti': '/sort/views',
        'cliccati': '/sort/clicks',
        'commentati': '/sort/comments',
        'argomentati': '/sort/threads',
        'votati': '/sort/votes',
        'più graditi': '/sort/top-rated',
        'meno graditi': '/sort/low-rated'}

#Numero di pagine    Inserire -1 per non mettere nessun limite
max_page = 3
#Pagina corrente
current_page = 1
#Pagina iniziale
start_page = 2
#Articoli totali estratti
tot_articles = 0

#Apertura foglio excel
wb = Workbook()
ws = wb.active

#Opzioni per il webdriver
options = EdgeOptions()
options.add_argument("start-maximized")
options.add_experimental_option('excludeSwitches', ['enable-logging'])
current_page = start_page

#Creazione webdriver           In questo caso Edge
driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()),options=options)

def analisi(query="",filter=""):
    global max_page,current_page,tot_articles,wb,ws,driver
    if current_page <= max_page:
        #Caricamento pagina
        query = 'search/query/' + query.replace(' ', '+')
        pageQuery = ""
        if(current_page > 1):
            pageQuery = "/pag/"+str(current_page)

        driver.get("https://www.firenzetoday.it/" + str(query) + filters[filter] + pageQuery)

        #Ricerca
        articles = driver.find_elements(By.CSS_SELECTOR, 'article')

        for article in articles:
            
            titleTxt = ""
            descriptionTxt = ""
            linkTxt = ""
            dateTxt = ""
            textTxt = ""

            try:
                title = article.find_element(By.CLASS_NAME, 'u-heading-04.u-mb-small.u-mt-none')
                titleTxt = title.text
            except:
                pass
            try:
                description = article.find_element(By.CLASS_NAME, 'u-body-04.u-color-secondary.u-mb-small')
                descriptionTxt = description.text
            except:
                pass

            main_window = driver.current_window_handle
            try:
                link = None
                try:
                    link = article.find_element(By.CLASS_NAME,'c-story__header').find_element(By.CSS_SELECTOR, 'a')
                except:
                    pass
                if (link):
                    linkTxt = link.get_attribute('href')
                    main_window = driver.current_window_handle
                    driver.execute_script("window.open('"+linkTxt+"', 'new_window')")
                    driver.switch_to.window(driver.window_handles[1])
                    sleep(1)
                    div = driver.find_element(By.CLASS_NAME, 'c-entry')
                    ps = div.find_elements(By.CSS_SELECTOR, 'p')
                    for p in ps:
                        textTxt += p.text
                    driver.close()
                    sleep(1)
                    driver.switch_to.window(main_window)
            except:
                driver.switch_to.window(main_window)
            driver.switch_to.window(main_window)
            try:
                date = article.find_element(By.CLASS_NAME, 'c-story__byline.u-label-08.u-color-secondary.u-mb-xsmall.u-block')
                dateTxt = date.text
            except:
                pass
            
            ws.append((titleTxt, descriptionTxt, dateTxt, linkTxt, textTxt))
            tot_articles+=1
            wb.save('FirenzeToday.xlsx')
        print("Pagina:",current_page," | ","Articoli trovati: " + str(len(articles))+" | " + "Articoli salvati: " + str(tot_articles))

        current_page +=1

        analisi(query,filter)



# Funzione utilizzata nel caso lo script venga interrotto con CTRL+C.
def save():
    global tot_articles,wb,driver
    print("Scraping interrotto \n" + str(tot_articles) + " articoli salvati \nAttendere terminazione script.")
    wb.save('FirenzeToday.xlsx')
    driver.quit()


try:
    #Signal per identificare quando lo script viene fermato con CTRL+C
    signal.signal(signal.SIGINT, signal.default_int_handler)
    analisi(query="Uffizi", filter="cliccati")
except KeyboardInterrupt:
    os._exit(0)

