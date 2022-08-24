import requests as req
import json
import xlsxwriter

reqst = req.get("https://restcountries.com/v3.1/all") #requisicao http
all = json.loads(reqst.content)

dic = {} #dicionario usado para padronizar e acessar os dados

for i in all:
    if 'capital' in i: #Confere se o país tem uma capital, senão atribui capital: '-'
        if 'currencies' in i: #Confere se o país tem moeda, senão atribui moeda: '-'
            dic[i['name']['common']] = {'capital': i['capital'][0], 'area': (i['area']), 'currencie': ','.join(list(i['currencies'].keys()))}
        else:
            dic[i['name']['common']] = {'capital': i['capital'][0], 'area': (i['area']), 'currencie': '-'}
    else:
        if 'currencies' in i:
            dic[i['name']['common']] = {'capital': '-', 'area': (i['area']), 'currencie': ','.join(list(i['currencies'].keys()))}
        else:
            dic[i['name']['common']] = {'capital': '-', 'area': (i['area']), 'currencie': '-'}


workbook = xlsxwriter.Workbook('planilha.xlsx') #cria arquivo xlsx
worksheet = workbook.add_worksheet()

title_style = workbook.add_format({'bold': 1, 'font_color': '#4F4F4F', 'font_size': 16,}) #estilo do titulo
title_style.set_center_across()
col_style = workbook.add_format({'bold': 1, 'font_color': '#808080', 'font_size': 12}) #estilo da legenda por coluna 
area_style = workbook.add_format({'num_format': '#,##0.00', 'align': 'right'}) #estilo da area

worksheet.write(0, 0,'Countries List', title_style)
worksheet.write_blank(0, 1, '', title_style)
worksheet.write_blank(0, 2, '', title_style)
worksheet.write_blank(0, 3, '', title_style)

col_title = ['Name', 'Capital', 'Area', 'Currencies']
row = 1
col = 0

for i in col_title: 
    worksheet.write(row, col, i, col_style)
    col += 1

col = 0
for i in sorted(dic.keys()): 
    worksheet.write(row, col, i)
    worksheet.write(row, col+1, dic[i]['capital'])
    worksheet.write(row, col+2, dic[i]['area'], area_style)
    worksheet.write(row, col+3, dic[i]['currencie'])
    row += 1

workbook.close()