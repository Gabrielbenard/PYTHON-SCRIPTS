#OBJETIVO É PEGA EM UMA PLANILHA OS DADOS DOS CEPS, IR NO BUSCACEP COLETAR TODOS OS DADOS E JOGAR NA PLANILHA NA FOLHA "Dados"

#biblioteca SELENIUM
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

#Biblioteca Pyautogui
import pyautogui as pa

#Biblioteca Openpyxl
from openpyxl import load_workbook

#pegamos o caminho + nome do arquivo no computador
arquivoCeps = 'C:\\Users\\GBERNARDINO\\PycharmProjects\\AutomationRPA\\PegaCEPS.xlsx'
planilhaDadosCeps = load_workbook(arquivoCeps)

#LEndo a sheet CEPA

FolhaSelecionada = planilhaDadosCeps["CEP"]

BrowserEdge = webdriver.Edge()
BrowserEdge.get("https://buscacepinter.correios.com.br/app/endereco/index.php")

BrowserEdge.find_element(By.XPATH,'//*[@id="endereco"]').send_keys("50720190")

BrowserEdge.find_element(By.XPATH,'//*[@id="endereco"]').send_keys(Keys.RETURN)

pa.sleep(4)

#Para cada linha até ultima linha
for linha in range(2,len(FolhaSelecionada["A"]) + 1):


    # clicar no nova Busca
    BrowserEdge.find_element(By.XPATH, '//*[@id="btn_nbusca"]').click()

    pa.sleep(3)

    #CEP da planilha e colocando na variável
    cepPesquisa = FolhaSelecionada["A%s" % linha].value


    BrowserEdge.find_element(By.XPATH, '//*[@id="endereco"]').send_keys(cepPesquisa)

    pa.sleep(3)

    BrowserEdge.find_element(By.XPATH, '//*[@id="btn_pesquisar"]').click()

    pa.sleep(3)

    Rua = BrowserEdge.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[1]').text
    Bairro = BrowserEdge.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[2]').text
    Localidade = BrowserEdge.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[3]').text
    print(Rua)
    print(Bairro)
    print(Localidade)

    #Seleciona a Sheet de Dados
    sheet_Dados_Imprimir_endereço = planilhaDadosCeps["Dados"]

    LinhaCorrentPlanilhaCEP = len(sheet_Dados_Imprimir_endereço["A"]) + 1

    #criando as variáveis de coluna para cada linha
    ColunaA = "A" + str(LinhaCorrentPlanilhaCEP)
    ColunaB = "B" + str(LinhaCorrentPlanilhaCEP)
    ColunaC = "C" + str(LinhaCorrentPlanilhaCEP)
    ColunaD = "D" + str(LinhaCorrentPlanilhaCEP)

    pa.sleep(5)

    #Imprimir as informações do site na planilha
    sheet_Dados_Imprimir_endereço[ColunaA] = Rua
    sheet_Dados_Imprimir_endereço[ColunaB] = Bairro
    sheet_Dados_Imprimir_endereço[ColunaC] = Localidade
    sheet_Dados_Imprimir_endereço[ColunaD] = cepPesquisa

    pa.sleep(3)

#Salvando o arquivo excel com novas informações
planilhaDadosCeps.save(filename=arquivoCeps)

# RUA = BrowserEdge.find_element(By.XPATH,'//*[@id="resultado-DNEC"]/tbody/tr/td[1]').text
# Bairro = BrowserEdge.find_element(By.XPATH,'//*[@id="resultado-DNEC"]/tbody/tr/td[2]').text
# Localidade = BrowserEdge.find_element(By.XPATH,'//*[@id="resultado-DNEC"]/tbody/tr/td[3]').text