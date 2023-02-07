from time import sleep
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ImageEX
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import datetime,timedelta
import urllib

def getTime(timeZ):

    timeZ = timeZ.replace("T"," ").replace("Z","")
    date = datetime.strptime(timeZ, '%Y-%m-%d %H:%M:%S')
    return (date + timedelta(hours=1))
    
wb = Workbook()
ws = wb.active

driver = webdriver.Edge('msedgedriver.exe')
driver.get("https://news.google.com/home?hl=it&gl=IT&ceid=IT:it")

driver.find_element(By.CLASS_NAME, value="VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.DuMIQc.LQeN7.Nc7WLe").click()
search = driver.find_element(By.CLASS_NAME, value="Ax4B8.ZAGvjd")
search.send_keys("italia")
search.send_keys(Keys.ENTER)

sleep(3)

container = driver.find_element(By.CLASS_NAME, value="lBwEZb.BL5WZb.GndZbb")

#CLASS NAME article MQsxIb.xTewfe.R7GTQ.keNKEd.j7vNaf.Cc0Z5d.EjqUne
#CLASS NAME title ipQwMb.ekueJc.RD0gLb
#CLASS NAME time WW6dff.uQIVzc.Sksgp.slhocf
#CLASS NAME from text wEwyrc.AVN2gc.WfKKme  
#CLASS NAME from image tvs3Id.tvs3Id.lqNvvd.ICvKtf.WfKKme.b1F67d
#CLASS NAME link VDXfz


# images = driver.find_elements(By.CLASS_NAME, value="tvs3Id.tvs3Id.lqNvvd.ICvKtf.WfKKme.b1F67d")
# print(images[0].get_attribute("src"))

#urllib.request.urlretrieve(str(images[0].get_attribute("src")),"Images/image{}.webp".format(0))
# data = urllib.request.urlopen(images[0].get_attribute("src")).read()
# file = open("image.webp", "wb")
# file.write(data)
# file.close()

# from PIL import Image
# im = Image.open("image.webp").convert("RGB")
# im.save("image.png","png")

# img = ImageEX("image.png")
# img.anchor = 'A1'

# ws.add_image(img)


# wb.save("Scraping.xlsx")

height1 = driver.execute_script("return document.body.scrollHeight")
for scrol in range(100,height1,100):
    driver.execute_script(f"window.scrollTo(0,{scrol})")
sleep(7)


articoli = container.find_elements(By.CLASS_NAME, value="MQsxIb.xTewfe.R7GTQ.keNKEd.j7vNaf.Cc0Z5d.EjqUne")

print(len(articoli))

for articolo in articoli:
    print(articolo.find_element(By.CLASS_NAME, value="ipQwMb.ekueJc.RD0gLb").text)
    print(getTime(articolo.find_element(By.CLASS_NAME, value="WW6dff.uQIVzc.Sksgp.slhocf").get_attribute("datetime")))
    from_text = articolo.find_element(By.CLASS_NAME, value="wsLqz.RD0gLb").text

    if from_text is None and from_text == "" and from_text.count()<1:
        print(articolo.find_element(By.CLASS_NAME, value="tvs3Id.tvs3Id.lqNvvd.ICvKtf.WfKKme.b1F67d").get_attribute("src"))
    else:
        print(from_text)

    print(articolo.find_element(By.CLASS_NAME, value="VDXfz").get_attribute("href"))
    print("\n\n")



sleep(50)