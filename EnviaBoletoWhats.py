#Import referente ao windows 7 gui
import win32con as win32con
from win32 import win32gui


#Import referente ao Selenium e Webdriver
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

#from webdriver_manager.chrome import ChromeDriverManager
#Outros imports
import time
import pandas as pd
import logging
import os


#caps = webdriver.DesiredCapabilities.CHROME.copy()
#caps['acceptInsecureCerts'] = True
#driver = webdriver.Chrome()

################### LOG para capturar erro no envio #######################
logging.basicConfig(filename='log.log', level=logging.DEBUG,format='%(asctime)s %(levelname)s %(funcName)s => %(message)s')


######Copia o perfil do chrome para o chromiun
dir_path = os.getcwd()
# O caminho do chromedriver
web_driver = os.path.join(dir_path, "chromedriver.exe")
# Caminho onde será criada pasta profile
profile = os.path.join(dir_path, "profile", "wpp")

options = webdriver.ChromeOptions()
# Configurando a pasta profile, para mantermos os dados da seção
options.add_argument(r"user-data-dir={}".format(profile))
# Inicializa o webdriver
web_driver = webdriver.Chrome(web_driver, options=options)

#número
#variavel para armazenar o número do telefone destino
#padrão 55 + ddd 2 digitos + numero
vNumeroTelefone = '55xxxxxxxx'
### Abre o whatsappweb
web_driver.get("http://wa.me/%s", vNumeroTelefone)
# Aguarda alguns segundos para validação manual do QrCode
web_driver.implicitly_wait(15)

element = web_driver.find_element_by_xpath('//*[@id="action-button"]')
web_driver.execute_script("arguments[0].click();", element)


element = web_driver.find_element_by_xpath('//*[@id="fallback_block"]/div/div/a')
web_driver.execute_script("arguments[0].click();", element)

web_driver.implicitly_wait(100)   # Segundos (implicitly_wait aguarda até que uma requisição seja feita ou os 100 segundos passem)


#acima ou então abrir o link e iniciar uma conversa com um contato ... a desenvolver
#web_driver.get("https://web.whatsapp.com/")



#funconando ... envia msg texto
vMsg = "Segue se boleto para pagamento ..."
web_driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]').send_keys(vMsg)

element = web_driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button')
web_driver.execute_script("arguments[0].click();", element)


#envia anexo
# Pressiona o botão de anexo
button = web_driver.find_elements_by_xpath("//*[@id='main']/footer/div/div/div[2]/div/div/span")
button[0].click()
web_driver.implicitly_wait(3)

# Pressiona o botão de documento
inp_xpath = "//*[@id='main']/footer/div/div/div[2]/div/span/div/div/ul/li[3]/button/span"
button = web_driver.find_elements_by_xpath(inp_xpath)
button[0].click()
time.sleep(10)   # segundo. Esta pausa é necessária


hdlg = 0
while hdlg == 0:
    hdlg = win32gui.FindWindow(None, "Abrir")

time.sleep(1)   # second. This pause is needed

# Set filename and press Enter key
hwnd = win32gui.FindWindowEx(hdlg, 0, "ComboBoxEx32", None)
hwnd = win32gui.FindWindowEx(hwnd, 0, "ComboBox", None)
hwnd = win32gui.FindWindowEx(hwnd, 0, "Edit", None)

filename = 'boleto.txt'
win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, None, filename)

# Press Save button
hwnd = win32gui.FindWindowEx(hdlg, 0, "Button", "&Abrir")

win32gui.SendMessage(hwnd, win32con.BM_CLICK, None, None)


element = web_driver.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div')
web_driver.execute_script("arguments[0].click();", element)

web_driver.implicitly_wait(100)
