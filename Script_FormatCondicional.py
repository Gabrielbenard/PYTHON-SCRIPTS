import xlsxwriter as xlw
import os

#SCRIPT PARA FORMATAÇÃO CONDICIONAL EM PLANILHA EXCEL UTILIZANDO A BIBLIOTECA XLSXWRITTER, >=50 VERDE, <50 VERMELHO

nomeCaminhoArquivo= 'C:\\Users\\GBERNARDINO\\PycharmProjects\\AutomationRPA\\FormataçãoCondicional.xlsx'

#Cria arquivo excel
planilhaExcel= xlw.Workbook(nomeCaminhoArquivo)

sheetDados= planilhaExcel.add_worksheet()

#adiciona Formato cor de fundo verde e fonte branca
formatoMaior = planilhaExcel.add_format({'bg_color':'green',
                                         'font_color': 'white',
                                         })

#adiciona Formato cor de fundo white e fonte branca
formatoMenor = planilhaExcel.add_format({'bg_color':'red',
                                         'font_color': 'white',
                                         })

#inserindo Tabela
inserirDados = [
     ['Coluna1','Coluna2', "Coluna 3","Coluna4"],
    [30,70,12,8],
    [23,245,10,81],
    [29,58,73,19],
]


sheetDados.write("A1",">50 estão em verde e <50 estão em vermelho")

for linha,range in enumerate(inserirDados):
    sheetDados.write_row(linha + 2, 1, range) # 'linha+2 '=  Linha 3, '1' coluna 2


#formatação condicional
sheetDados.conditional_format("B4:E6",{"type": 'cell',
                                       "criteria": ">=",
                                       'value': 50,
                                       "format": formatoMaior})

#formatação condicional
sheetDados.conditional_format("B4:E6",{"type": 'cell',
                                       "criteria": "<",
                                       'value': 50,
                                       "format": formatoMenor})


#Finaliza ações no workbook
planilhaExcel.close()

os.startfile(nomeCaminhoArquivo)