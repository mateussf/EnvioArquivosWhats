#Import referente ao windows 7 gui
import win32con as win32con
from win32 import win32gui


#Import referente ao Selenium e Webdriver
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
#from selenium.webdriver.support.ui import WebDriverWait

#from webdriver_manager.chrome import ChromeDriverManager
#Outros imports
import time
import pandas as pd
import logging
import os


def IniciaBrowser():
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
    #maximizar janela
    web_driver.maximize_window()
    
    return web_driver

def NavegaAteConversa():
    web_driver.implicitly_wait(15)

    element = web_driver.find_element_by_xpath('//*[@id="action-button"]')
    web_driver.execute_script("arguments[0].click();", element)


    element = web_driver.find_element_by_xpath('//*[@id="fallback_block"]/div/div/a')
    web_driver.execute_script("arguments[0].click();", element)

    

def AbreArquivo():
    vArquivo = open('arquivos.txt', 'r')

    #arquivo = open('arquivo.txt', 'r')
    lista = vArquivo.readlines() # readlinesssssss

    vContatos = []
    vArquivos = []
    for vLinha in lista:
        vLinhax = (vLinha.split(';'))
        vContatos.append(vLinhax[0])
        vArquivos.append(vLinhax[1])

    return vContatos, vArquivos

    vArquivo.close()

def EnviaMensagem(vMensagem):
    web_driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]').send_keys(vMensagem)
    #web_driver.send_keys(Keys.RETURN)
    element = web_driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button').send_keys(Keys.RETURN)

def abre_arquivo(dir):
    # Loop até que a caixa de diálogo Open seja exibida
    hdlg = 0
    while hdlg == 0:
        hdlg = win32gui.FindWindow(None, "Abrir")

    
    # Define o nome do arquivo e pressione a tecla Enter
    hwnd = win32gui.FindWindowEx(hdlg, 0, "ComboBoxEx32", None)
    hwnd = win32gui.FindWindowEx(hwnd, 0, "ComboBox", None)
    hwnd = win32gui.FindWindowEx(hwnd, 0, "Edit", None)
    win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, None, dir)
    # Pressiona o botão Salvar
    hwnd = win32gui.FindWindowEx(hdlg, 0, "Button", "&Abrir")
    win32gui.SendMessage(hwnd, win32con.BM_CLICK, None, None)

vContatos, vArquivos = AbreArquivo()


for vContato in vContatos:
    ### Abre o whatsappweb
    web_driver = IniciaBrowser()
    vLink = ("http://wa.me/" +  vContato)
    
    web_driver.get(vLink)
    #WebDriverWait(web_driver).until(document_initialised)
    
    # Aguarda alguns segundos para validação manual do QrCode
    web_driver.implicitly_wait(15)


    NavegaAteConversa() 

    #envia mensagem antes do anexo
    EnviaMensagem("Olá, segue o arquivo ..")
    
    web_driver.implicitly_wait(100)   # Segundos (implicitly_wait aguarda até que uma requisição seja feita ou os 100 segundos passem)


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


    I = vContatos.index(vContato)
    vCaminho = os.getcwd()+'\\arquivos'

    abre_arquivo(vCaminho)
    time.sleep(1)
    #Abre diretório de boletos (inserido caso o primeiro envio da função não execute)
    abre_arquivo(vCaminho)
    #Abre Arquivo nome do cliente . pdf
    vArquivo = vArquivos[I].rstrip()
    abre_arquivo(f'{vArquivo}')


    send = web_driver.find_element_by_xpath(f"//span[@data-icon='send']")
    send.click()
    time.sleep(2)
   
    web_driver.implicitly_wait(100)

    web_driver.quit()
