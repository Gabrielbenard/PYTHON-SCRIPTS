import xlsxwriter as xlw
import os

nomeCaminhoArquivo= 'C:\\Users\\GBERNARDINO\\PycharmProjects\\AutomationRPA\\PrimeiroExemplo.xlsx'

#Cria arquivo do zero excel
workbook= xlw.Workbook(nomeCaminhoArquivo)

sheetPadrao= workbook.add_worksheet()

#cor de fundo da célula
# corfundoAmarela = workbook.add_format({'fg_color':'yellow'})


#Cor da fonte da célula
corFonteAzul = workbook.add_format()
corFonteAzul.set_font_color("blue")

corfundoAmarela = workbook.add_format({'align':'center',
                                       'font_color':'green',
                                        'bold': 'true',
                                        'bg_color': 'black'
                                       })

sheetPadrao.write("A1","Nome",corfundoAmarela)
sheetPadrao.write("B1","Idade",corfundoAmarela)
sheetPadrao.write("A2","Amanda",corFonteAzul)
sheetPadrao.write("B2",22,corFonteAzul)
sheetPadrao.write("A3","José",corFonteAzul)
sheetPadrao.write("B3",29,corFonteAzul)


#Finaliza ações no workbook
workbook.close()

os.startfile(nomeCaminhoArquivo)