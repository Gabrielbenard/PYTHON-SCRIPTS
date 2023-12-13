import xlsxwriter
import os


caminhoArquivo= 'C:\\Users\\GBERNARDINO\\PycharmProjects\\AutomationRPA\\TesteXlsxwritter.xlsx'


#Abrindo o Workbook em segundo plano
planilhaCriada= xlsxwriter.Workbook(caminhoArquivo)


Folha1= planilhaCriada.add_worksheet("folha1")


#Escrevendo na celula em backgrounda
Folha1.write("A1","Nome")

#fechando arquivo em background
planilhaCriada.close()

#Abrindo em foreground
os.startfile(caminhoArquivo)
