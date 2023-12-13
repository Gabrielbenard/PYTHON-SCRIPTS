#SELENIUM
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

#Biblioteca Pyautogui
import pyautogui as pa

#Biblioteca Openpyxl
from openpyxl import load_workbook


#pegamos o caminho + nome do arquivo no computador, também abrimos a planilha  e por Fim a folha
ArquivoDadosExcel ="C:\\Users\\GBERNARDINO\\PycharmProjects\\AutomationRPA\\DadosPessoais.xlsx"
woorkbookDadosExcel = load_workbook(ArquivoDadosExcel)
FolhaDados=woorkbookDadosExcel["Dados"]

for linha in range(2,len(FolhaDados["A"])+1):

    # Abrindo o web formulário
    BrowserSurvey = webdriver.Edge()
    BrowserSurvey.get('https://pt.surveymonkey.com/r/XRZ9HPP')

    #Ler e coletar os dados a linha
    Nome = FolhaDados["A%s" % linha].value
    ID = FolhaDados["B%s" % linha].value
    Telefone = FolhaDados["C%s" % linha].value
    Email = FolhaDados["D%s" % linha].value
    DataNascimento = str(FolhaDados["E%s" % linha].value)[0:10]
    Sexo = FolhaDados["F%s" % linha].value

    pa.sleep(3)

    #preencher formulário web
    BrowserSurvey.find_element(By.XPATH,'//*[@id="162685332"]').send_keys(Nome)   #nome
    pa.sleep(1)
    BrowserSurvey.find_element(By.XPATH, '//*[@id="162685494"]').send_keys(ID)  #ID
    pa.sleep(1)
    BrowserSurvey.find_element(By.XPATH, '//*[@id="162685541"]').send_keys(Telefone) #Telefone
    pa.sleep(1)
    BrowserSurvey.find_element(By.XPATH, '//*[@id="162685567"]').send_keys(Email)  #Email
    pa.sleep(1)
    BrowserSurvey.find_element(By.XPATH, '//*[@id="162685619"]').send_keys(DataNascimento)  # Data de nascimento

    if str(Sexo) == "Masculino":
        BrowserSurvey.find_element(By.XPATH, '//*[@id="162686044_1191113511_label"]').click()   #radio button
    else:
        BrowserSurvey.find_element(By.XPATH, '//*[@id="162686044_1191113512_label"]').click()   #radio button

    pa.sleep(1)

    #Concluído e eviar
    BrowserSurvey.find_element(By.XPATH, '//*[@id="patas"]/main/article/section/form/div[2]/button').click() # Concluido

    pa.sleep(2)
    #Fechando o navegador
    BrowserSurvey.close()

    pa.sleep(2)
