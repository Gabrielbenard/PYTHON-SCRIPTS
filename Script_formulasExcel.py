import xlsxwriter as xlw
import os

nomeCaminhoArquivo= 'C:\\Users\\GBERNARDINO\\PycharmProjects\\AutomationRPA\\FormulasExemplo.xlsx'

#Cria arquivo excel
workbook= xlw.Workbook(nomeCaminhoArquivo)

sheetPadrao= workbook.add_worksheet()


sheetPadrao.write("A1",8)
sheetPadrao.write("A2",15)
sheetPadrao.write("A3",18)
sheetPadrao.write("A4",20)
sheetPadrao.write("A5",4)
sheetPadrao.write("A5",'Gabriel')
sheetPadrao.write("A7",'Gama')

sheetPadrao.write("B1",8)
sheetPadrao.write("B2",17)
sheetPadrao.write("B3",48)
sheetPadrao.write("B4",20)
sheetPadrao.write("B5",1851)
sheetPadrao.write("B6",'Bernardino')
sheetPadrao.write("B7",'Louco')

sheetPadrao.write_formula("C1","=A1+B1")
sheetPadrao.write_formula("C2","=A2-B2")
sheetPadrao.write_formula("C3","=A3/B3")
sheetPadrao.write_formula("C4","=A4*B4")
sheetPadrao.write_formula("C5",'=CONCATENATE(A5,"  ",B6)') # OBS: utilizando aspas duplas e nos espaços entre celulas também não irá funciona

sheetPadrao.merge_range("A8:B8","Merge Range")



#Finaliza ações no workbook
workbook.close()

os.startfile(nomeCaminhoArquivo)
