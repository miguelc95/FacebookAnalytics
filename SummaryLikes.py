import urllib
import json
import sys
import os
import facepy
import xlrd
import xlwt
from xlutils.copy import copy

#Miguel Cuellar



#AccessToken y version de el api GRAPH
graph = facepy.GraphAPI('YOUR-ACCESS-TOKEN-HERE',version='2.9')

#Datos del libro de excel a leer para sacar los datos
book = xlrd.open_workbook('WORKBOOK-NAME.xlsx')
max_nb_row = 0
wb = xlwt.Workbook()
ws = wb.add_sheet('Datos de posts')
#iterar por las hojas del libro
for sheet in book.sheets():
  max_nb_row = max(max_nb_row, sheet.nrows)

#iterar por los renglones del libro
for row in range(max_nb_row) :
  for sheet in book.sheets() :
    if row < sheet.nrows :
        #leer los ids de los posts a analizar
        if str(sheet.cell(row, 1).value) != 'id' :
            createdTime = str(sheet.cell(row, 0).value)
            postID = str(sheet.cell(row, 1).value)
            message = (sheet.cell(row, 2).value)
            #Hacer el get del api con los parametros que se necesitan
            response = graph.get(postID+'?fields=shares,status_type,likes.limit(1).summary(true),comments.limit(1).summary(true),reactions.type(LIKE).limit(0).summary(1).as(like),reactions.type(HAHA).limit(0).summary(1).as(haha),reactions.type(WOW).limit(0).summary(1).as(wow),reactions.type(SAD).limit(0).summary(1).as(sad),reactions.type(LOVE).limit(0).summary(1).as(love),reactions.type(ANGRY).limit(0).summary(1).as(angry)')

            #Verificar si el post es compartido
            if 'shares' in response:
                ws.write(row, 0, createdTime)
                ws.write(row, 1, postID)
                ws.write(row, 2, message)
                ws.write(row, 3, str(response['likes']['summary']['total_count']))
                ws.write(row, 4, str(response['comments']['summary']['total_count']))
                ws.write(row, 5, str(response['wow']['summary']['total_count']))
                ws.write(row, 6, str(response['sad']['summary']['total_count']))
                ws.write(row, 7, str(response['angry']['summary']['total_count']))
                ws.write(row, 8, str(response['love']['summary']['total_count']))
                ws.write(row, 9, str(response['haha']['summary']['total_count']))
                ws.write(row, 10, str(response['shares']['count'])) #Si no es compartido se pone el share count
                if 'status_type' in response:
                    ws.write(row, 11, str(response['status_type']))


            else:
                ws.write(row, 0, createdTime)
                ws.write(row, 1, postID)
                ws.write(row, 2, message)
                ws.write(row, 3, str(response['likes']['summary']['total_count']))
                ws.write(row, 4, str(response['comments']['summary']['total_count']))
                ws.write(row, 5, str(response['wow']['summary']['total_count']))
                ws.write(row, 6, str(response['sad']['summary']['total_count']))
                ws.write(row, 7, str(response['angry']['summary']['total_count']))
                ws.write(row, 8, str(response['love']['summary']['total_count']))
                ws.write(row, 9, str(response['haha']['summary']['total_count']))
                #Si no es compartido no se pone el share count
                if 'status_type' in response:
                    ws.write(row, 11, str(response['status_type']))


#Se graba el libro de excel con los resultados
wb.save('OUTPUT-WORBOOK_NAME.xls')
