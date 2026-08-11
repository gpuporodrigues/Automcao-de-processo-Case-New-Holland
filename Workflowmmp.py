#Bibliotecas
from operator import index
import time
import pyautogui
import selenium
import playwright   
import pandas as pd
import openpyxl
import numpy

#importando funçoes das bibliotecas
from selenium import webdriver
driver = webdriver.Edge()
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
#_____________________________________________________________________
#Acesso Workflow(login e senha+click acessar)
url_Workflow = 'https://mmp.cnhindustrial.ind.br/#/login'
driver.get(url_Workflow)
driver.maximize_window()
time.sleep(2)
driver.find_element('id','mat-input-0').send_keys('!Confidencial!')
driver.find_element('id','mat-input-1').send_keys('!Confidencial!')
pyautogui.press('enter')
time.sleep(10)
#_____________________________________________________________________
#MMP
requisicaommp = driver.find_element(
    By.XPATH,
    "//span[text()='Requisição']"
)
requisicaommp.click()

requisicaompp0 = driver.find_element(
    By.XPATH,
    "//button[@role='menuitem']"
)
requisicaompp0.click()

#adiciona a solicitação
adicionarmmp = driver.find_element(
    By.XPATH,
    "//button[.//mat-icon[contains(text(),'add_circle_outline')]]"
)
adicionarmmp.click()
time.sleep(2)
driver.find_element('name','plant').click()

cu01mmp = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable(
        (By.XPATH, "//mat-option//span[contains(text(),'CU01 - Curitiba')]")
    )
)
cu01mmp.click()

driver.find_element('name','item').click()

consumommp = driver.find_element(
    By.XPATH,
    "//mat-option[.//span[contains(text(),'CON - Consumo')]]"
)
consumommp.click()

driver.find_element('name','gama').click()

geralmmp = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable(
        (By.XPATH, "//mat-option[.//span[normalize-space()='Geral - Geral']]")
    )
)
geralmmp.click()

#importa dados do excel
dadosinterface = "interfacemmp.xlsx"
df = pd.read_excel(dadosinterface, header=None)
ccpep = str(df.iloc[0, 0])  # ajuste a posição conforme sua planilha

if ccpep[0] == "B":
   campo_cc = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.NAME, "centrocusto"))
 )
   campo_cc.send_keys(ccpep)
else:
    driver.find_element(
       By.XPATH,
       "//input[@formcontrolname='pep']"
   ).send_keys(ccpep)

addmateriais = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(., 'Adicionar Materiais')]")
    )
)
addmateriais.click()

cargamassiva = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable(
        (By.XPATH, "//button[normalize-space()='Carga Massiva']")
    )
)
cargamassiva.click()
arquivo = r"C:\Users\!Confidencial!\Desktop\Automacao\Massivelload.xlsx"

input_file = driver.find_element(
    By.XPATH,
    "//input[@type='file']"
)

driver.execute_script(
    "arguments[0].style.display='block';",
    input_file
)

input_file.send_keys(
    arquivo
)
time.sleep(2)
pyautogui.click(1106,712) 

time.sleep(60)
