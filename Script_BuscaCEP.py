from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pyautogui as pa
from selenium.webdriver.common.by import By
import xlsxwriter as xlw

CEP = input("Digite o CEP: ")


BrowserEdge = webdriver.Edge()

BrowserEdge.get("https://buscacepinter.correios.com.br/app/endereco/index.php")

BrowserEdge.find_element(By.XPATH,'//*[@id="endereco"]').send_keys(CEP)

BrowserEdge.find_element(By.XPATH,'//*[@id="endereco"]').send_keys(Keys.RETURN)

pa.sleep(4)


RUA = BrowserEdge.find_element(By.XPATH,'//*[@id="resultado-DNEC"]/tbody/tr/td[1]').text
Bairro = BrowserEdge.find_element(By.XPATH,'//*[@id="resultado-DNEC"]/tbody/tr/td[2]').text
Localidade = BrowserEdge.find_element(By.XPATH,'//*[@id="resultado-DNEC"]/tbody/tr/td[3]').text

pathExcelBuscaCep = "C:\\Users\\GBERNARDINO\\PycharmProjects\\AutomationRPA\\BuscaCEP.xlsx"

#Abrindo planilha background
BackExcel = xlw.Workbook(pathExcelBuscaCep)

Folha1 = BackExcel.add_worksheet("CEP")

#Escrevendo titulos
Folha1.write("A1","CEP")
Folha1.write("B1","RUA")
Folha1.write("C1","Bairro")
Folha1.write("D1","Localidade")

#Escrevendo Valores
Folha1.write("A2",CEP)
Folha1.write("B2",RUA)
Folha1.write("C2",Bairro)
Folha1.write("D2",Localidade)


BackExcel.close()