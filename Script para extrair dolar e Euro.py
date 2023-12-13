from selenium import webdriver

from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.keys import Keys
import pyautogui as Stoptime
import xlsxwriter


#trabalhar  com as atualizações mais recentes
from selenium.webdriver.common.by import By


#Instanciando o navegador e acesso as configurações do Edge
driver = webdriver.Edge()
driver.get("https://www.google.com.br/")

Stoptime.sleep(4)


#encontrar o elemento pelo NAME UTILIZANDO ESPECIONAR E DÁ O SEU VALOR, E ENTÃO ESCREVER DOLAR HOJE e envia na pesquisa
driver.find_element(By.NAME, "q").send_keys("Dolar hoje")


Stoptime.sleep(4)

##encontrar o elemento pelo NAME UTILIZANDO ESPECIONAR E ENVIA NO ENTER
driver.find_element(By.NAME, "q").send_keys(Keys.RETURN)

Stoptime.sleep(4)

#Pegando o Xpath
valorDolarGoogle = driver.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').text

print(f"o valor do dolar hoje é {valorDolarGoogle} Reais")

#limpa o texto no localA
driver.find_element(By.NAME, "q").clear()

#encontrar o elemento pelo NAME UTILIZANDO ESPECIONAR E DÁ O SEU VALOR, E ENTÃO ESCREVER EURO HOJE e envia na pesquisa
driver.find_element(By.NAME, "q").send_keys("Euro hoje")

#encontrar o elemento pelo NAME UTILIZANDO ESPECIONAR E ENVIA NO ENTER
driver.find_element(By.NAME, "q").send_keys(Keys.RETURN)

Stoptime.sleep(4)

valorEuroGoogle = driver.find_element(By.XPATH,'//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').text

print(f"o valor do Euro hoje é {valorEuroGoogle} Reais")

Stoptime.sleep(4)

# 2) ------ ESCREVENDO NO EXCEL --------------


caminhoArquivo= 'C:\\Users\\GBERNARDINO\\PycharmProjects\\AutomationRPA\\TesteXlsxwritter.xlsx'


#Abrindo o Workbook em segundo plano
planilhaCriada= xlsxwriter.Workbook(caminhoArquivo)

#SOBRESCREEVE A PRIMEIRA FOLHA a
Folha1= planilhaCriada.add_worksheet("Dolar_Euro")


#Escrevendo na celula em background
Folha1.write("A1","Valor Dolar em Reais")
Folha1.write("A2",valorDolarGoogle)
Folha1.write("B1","Valor Euro em Reais")
Folha1.write("B2",valorEuroGoogle)

#fechando arquivo em background
planilhaCriada.close()

