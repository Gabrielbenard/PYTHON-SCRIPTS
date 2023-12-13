from openpyxl import load_workbook
import os

pathArquivo= 'C:\\Users\\GBERNARDINO\\PycharmProjects\\AutomationRPA\\DeletarLinhasColunas.xlsx'

planilhaAberta = load_workbook(pathArquivo)

sheetEscolhida = planilhaAberta["Dados"]

sheetEscolhida.delete_rows(3)
sheetEscolhida.delete_rows(3) #Quando deletar a linha 3 a linha 4 irá para a 3 novamente
#sheetEscolhida.detele_cols() #Deletar colunas

planilhaAberta.save(pathArquivo)

os.startfile(pathArquivo)