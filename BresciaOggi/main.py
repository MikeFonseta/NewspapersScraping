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

mesi = {'gennaio': '1',
        'febbraio': '2',
        'marzo': '3',
        'aprile': '4',
        'maggio': '5',
        'giugno': '6',
        'luglio': '7',
        'agosto': '8',
        'settembre': '9',
        'ottobre': '10',
        'novembre': '11',
        'dicembre': '12',}

#Numero di pagine    Inserire -1 per non mettere nessun limite
max_page = 2
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
current_page = starting_page

#Creazione webdriver           In questo caso Edge
driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()),options=options)

def analisi(query="",contained_words=""):
    global max_page,current_page,tot_articles,wb,ws,driver
    #Caricamento pagina
    pageQuery = ""
    if current_page > 1:
        pageQuery = "&page="+str(current_page)

    driver.get("https://www.bresciaoggi.it/altro/ricerca?q=" + str(query) + " " + ' '.join(f"'{w}'" for w in contained_words.split(" ")) + pageQuery)

    active_page = (int)(driver.find_element(By.CLASS_NAME, 'active').text)
    
    #Ricerca
    if (current_page <= max_page or max_page==-1) and active_page == current_page:
        articles = WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "article")))
        for article in articles:
            title = article.find_element(By.CSS_SELECTOR, 'h2').text
            link = article.find_element(By.CSS_SELECTOR, 'h2').find_element(By.CSS_SELECTOR, 'a').get_attribute('href')
            date = str(article.find_element(By.CLASS_NAME, 'date').text)
            dateSplit = date.split(' ')
            dateSplit[1] = mesi[dateSplit[1]]

            date = dateSplit[0] + "/" + dateSplit[1] + "/" + dateSplit[2]

            ws.append([title, datetime(int(dateSplit[2]), int(dateSplit[1]), int(dateSplit[0]), 0, 0).strftime('%d/%m/%Y') , link])
            tot_articles+=1
            wb.save('BresciaOggi.xlsx')
        
        print("Pagina:",current_page," | ","Articoli trovati: " + str(len(articles))+" | " + "Articoli salvati: " + str(tot_articles))

        #Carico prossima pagina
        current_page+=1
        analisi(query,contained_words)
    elif active_page != current_page and starting_page > 1:
        print("Ultima pagina per " + query + " " + ' '.join(f"'{w}'" for w in contained_words.split(" ")) + "è: " + str(active_page))
        print("Starting page inserita: " + str(starting_page))
        driver.quit()
    elif active_page != current_page:
        driver.quit()

# Funzione utilizzata nel caso lo script venga interrotto con CTRL+C.
def save():
    global tot_articles,wb,driver
    print("Scraping interrotto \n" + str(tot_articles) + " articoli salvati \nAttendere terminazione script.")
    wb.save('BresciaOggi.xlsx')
    driver.quit()


try:
    #Signal per identificare quando lo script viene fermato con CTRL+C
    signal.signal(signal.SIGINT, signal.default_int_handler)
    analisi(query="Cultura", contained_words="teatri Giovani")
except KeyboardInterrupt:
    os._exit(0)

